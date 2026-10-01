'use strict';

const MAX_COL = 16384; // XFD
const MAX_ROW = 1048576;

// Structural, workbook-aware rewriting of A1-style formulas when rows are
// inserted into (or removed from) a single worksheet.
//
// Unlike slideFormula (used for shared formulas), absolute row markers
// ($n) do not pin a reference. The point of a structural edit is that the
// identity of every cell on or after the edit point moves to its new row.
//
// Reference grammar recognised:
//   [sheet-prefix!] endpoint[:endpoint]
// sheet-prefix:
//   'Sheet Name'   (single quotes doubled inside)
//   or an unquoted run of non-operator characters
// endpoint:
//   $?COL$?ROW   e.g. A1, $A$1, A$1
//   $?COL:$?COL  whole columns (rows unaffected)
//   $?ROW:$?ROW  whole rows
// A bare number is never a reference.
//
// Text inside a double-quoted string literal is left untouched.
// Constructs that cannot be maintained reliably (INDIRECT, external
// workbook links, 3D references spanning the edited sheet) raise an
// explicit error instead of silently pointing at wrong cells.

// A run that may form an unquoted sheet prefix: anything that contains no
// whitespace, quotes, bang separator or formula/operator characters.
// Sheet names containing spaces are always quoted in formulas.
const UNQUOTED_SHEET = '[^\\s!\'"\\\\/?*\\[\\]()+,;@^~%=&<>\\-:#]';

const COL = '\\$?[A-Za-z]{1,3}';
const ROW = '\\$?[1-9][0-9]{0,6}';
const ENDPOINT = `(?:${COL}(?:${ROW})?|${ROW})`;
const REFERENCE = new RegExp(
  `(?:('(?:[^']|'')*'|${UNQUOTED_SHEET}+)!)?(?:(${ENDPOINT}):(${ENDPOINT})|(${COL}(?:${ROW})?))`,
  'g'
);
const THREE_D = new RegExp(
  `^('(?:[^']|'')*'|${UNQUOTED_SHEET}+):('(?:[^']|'')*'|${UNQUOTED_SHEET}+)!`
);
const TABLE_REF = /^[A-Za-z_\\][A-Za-z0-9_.]*\s*\[/;

function colToNumber(letters) {
  let n = 0;
  const upper = letters.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    n = (n * 26) + (upper.charCodeAt(i) - 64);
  }
  return n;
}

function parseEndpoint(ep) {
  const m = /^(\$?)([A-Za-z]{1,3})?(\$?)([1-9][0-9]{0,6})?$/.exec(ep);
  return {
    colDollar: m[1] || '',
    col: m[2] || '',
    rowDollar: m[3] || '',
    row: m[4] ? parseInt(m[4], 10) : undefined,
  };
}

function encodeEndpoint(p) {
  return `${p.colDollar}${p.col}${p.rowDollar}${p.row === undefined ? '' : p.row}`;
}

function isValidEndpoint(ep) {
  const p = parseEndpoint(ep);
  if (!p.col && p.row === undefined) return false;
  if (p.col && colToNumber(p.col) > MAX_COL) return false;
  if (p.row !== undefined && p.row > MAX_ROW) return false;
  return true;
}

function endpointKind(ep) {
  if (/^(\$?)([1-9][0-9]{0,6})$/.test(ep)) return 'row';
  if (/^(\$?)([A-Za-z]{1,3})$/.test(ep)) return 'col';
  return 'cell';
}

function sheetFromPrefix(prefix) {
  if (prefix.startsWith('\'')) {
    return prefix.slice(1, -1).replace(/''/g, '\'');
  }
  return prefix;
}

function isBoundary(ch) {
  // '.' is deliberately excluded: dotted names (FOO.B2 style tokens are
  // part of a name) must not be treated as references. '"' can follow a
  // closed string literal, after which another reference is legal.
  return ch === undefined || /[ \t\r\n()+\-*/&<>=%,:;{}!~^@#"]/.test(ch);
}

// Right-hand boundary: a cell reference can never be followed by '('
// (that makes LOG10(, FOO( a function/name call).
function isRefEnd(ch) {
  return ch !== '(' && isBoundary(ch);
}

// Map the row of a single cell identity through the splice.
function mapIdentityRow(row, start, count) {
  if (count > 0) {
    return row >= start ? row + count : row;
  }
  if (count < 0) {
    const deleted = -count;
    const end = start + deleted - 1;
    if (row < start) return row;
    if (row > end) return row - deleted;
    return undefined; // deleted
  }
  return row;
}

// Map a range endpoint row. A range crossed by an insertion keeps
// expressing the original cells plus the new rows (top stays, bottom
// expands); a range wholly on or after the edit moves with its data.
function mapRangeEndpoint(which, row, start, count, topRow) {
  if (count > 0) {
    return row >= start ? row + count : row;
  }
  if (count < 0) {
    const deleted = -count;
    const end = start + deleted - 1;
    if (row < start) return row;
    if (row > end) return row - deleted;
    return which === 'top' ? start : Math.max(start, topRow || start);
  }
  return row;
}

/**
 * Validate a formula before any mutation. Throws on constructs that
 * cannot be maintained reliably.
 */
function assertTransformable(formula, currentSheet) {
  if (/\bINDIRECT\s*\(/i.test(formula)) {
    throw new Error(
      `Cannot structurally edit workbook: formula "${formula}" on sheet "${currentSheet}" ` +
        'uses INDIRECT, whose references cannot be maintained'
    );
  }
  if (/\[[^\][]+\]/.test(formula)) {
    throw new Error(
      `Cannot structurally edit workbook: formula "${formula}" on sheet "${currentSheet}" ` +
        'uses a bracketed external/structured reference that cannot be maintained'
    );
  }
}

function resolveSheet(sheetNames, name) {
  if (!sheetNames) return {name, found: true};
  const found = sheetNames.find(s => s.toLowerCase() === name.toLowerCase());
  return found ? {name: found, found: true} : {name, found: false};
}

/**
 * Rewrite ordinary cell / shared-master / array formulas. Every
 * reference is treated as an identity (absolute markers irrelevant).
 *
 * @param {string} formula       formula text without leading '='
 * @param {{sheetName:string,start:number,count:number}} edit
 * @param {string} currentSheet  sheet that owns the formula
 * @param {string[]} [sheetNames] all sheet names (3D span resolution)
 * @returns {string} rewritten formula
 */
function transformFormula(formula, edit, currentSheet, sheetNames) {
  return rewrite(formula, {
    edit,
    currentSheet,
    sheetNames,
    singleIsIdentity: true,
    mapEndpoint: (p, isRangeEnd, topRow, referencedSheet) => {
      if (referencedSheet !== edit.sheetName) return null;
      if (isRangeEnd) {
        return mapRangeEndpoint('bottom', p.row, edit.start, edit.count, topRow);
      }
      return mapRangeEndpoint('top', p.row, edit.start, edit.count, topRow);
    },
  });
}

/**
 * Rewriter for the template formulas carried by data-validation and
 * conditional-formatting rules. Such a formula is stored once relative
 * to the top-left cell of the rule's sqref and evaluated per covered
 * cell.
 *
 *  - absolute row ($n): identity semantics - the referenced cell moved
 *  - relative row (n):  slides with the anchor cell by anchorDelta rows;
 *    when the insertion lands inside the rule range the anchor stays put
 *    (anchorDelta 0), so the template keeps evaluating against the new
 *    rows naturally.
 *
 * References to sheets other than the edited one are untouched.
 */
function transformRuleFormula(formula, edit, currentSheet, sheetNames, anchorDelta) {
  const crossSheet = currentSheet !== edit.sheetName;
  return rewrite(formula, {
    edit,
    currentSheet,
    sheetNames,
    singleIsIdentity: false,
    mapEndpoint: (p, _isRangeEnd, _topRow, referencedSheet) => {
      if (referencedSheet !== edit.sheetName) return null;
      // Cross-sheet references are always identities: there is no anchor
      // relationship between a rule on another sheet and the edited cells.
      if (crossSheet || p.rowDollar) {
        const mapped = mapIdentityRow(p.row, edit.start, edit.count);
        return mapped === undefined ? edit.start : mapped;
      }
      return Math.max(1, p.row + anchorDelta);
    },
  });
}

function rewrite(formula, ctx) {
  const {edit, currentSheet, sheetNames} = ctx;
  const replacements = [];
  const failures = [];

  let i = 0;
  const n = formula.length;

  while (i < n) {
    const ch = formula[i];

    // double-quoted string literal ("" is an escaped quote)
    if (ch === '"') {
      i++;
      while (i < n) {
        if (formula[i] === '"') {
          if (formula[i + 1] === '"') {
            i += 2;
          } else {
            i++;
            break;
          }
        } else {
          i++;
        }
      }
      continue;
    }

    // 3D reference: SheetA:SheetB!A1 or 'A B':'C D'!A1
    const threeD = THREE_D.exec(formula.substring(i));
    if (threeD && isBoundary(formula[i - 1])) {
      const span = {
        text: threeD[0],
        first: sheetFromPrefix(threeD[1]),
        last: sheetFromPrefix(threeD[2]),
      };
      if (threeDTouches(span, edit, sheetNames)) {
        throw new Error(
          `Cannot structurally edit workbook: formula "${formula}" on sheet "${currentSheet}" ` +
            `uses a 3D reference "${span.text}" spanning the edited sheet "${edit.sheetName}"`
        );
      }
      i += span.text.length;
      continue;
    }

    if (ch === '\'') {
      // could be a quoted sheet prefix starting a reference
      REFERENCE.lastIndex = i;
      const qm = REFERENCE.exec(formula);
      if (
        qm &&
        qm.index === i &&
        isBoundary(formula[i - 1]) &&
        isRefEnd(formula[i + qm[0].length])
      ) {
        applyMatch(qm, i);
        i += qm[0].length;
      } else {
        i = skipQuoted(formula, i);
      }
      continue;
    }

    // structured table reference: Name[...]
    const table = TABLE_REF.exec(formula.substring(i));
    if (table && isBoundary(formula[i - 1])) {
      let depth = 0;
      let j = i;
      while (j < n) {
        if (formula[j] === '[') depth++;
        else if (formula[j] === ']') {
          depth--;
          if (depth === 0) {
            j++;
            break;
          }
        }
        j++;
      }
      i = j;
      continue;
    }

    REFERENCE.lastIndex = i;
    const m = REFERENCE.exec(formula);
    if (
      m &&
      m.index === i &&
      isBoundary(formula[i - 1]) &&
      isRefEnd(formula[i + m[0].length])
    ) {
      applyMatch(m, i);
      i += m[0].length;
      continue;
    }

    i++;
  }

  if (failures.length) {
    throw new Error(
      `Cannot structurally edit workbook: formula "${formula}" on sheet "${currentSheet}" ${failures[0]}`
    );
  }

  let out = formula;
  for (let k = replacements.length - 1; k >= 0; k--) {
    const r = replacements[k];
    out = out.substring(0, r.start) + r.text + out.substring(r.end);
  }
  return out;

  function applyMatch(m, matchStart) {
    const token = m[0];
    const prefix = m[1];
    const refSheet = prefix
      ? resolveSheet(sheetNames, sheetFromPrefix(prefix))
      : {name: currentSheet, found: true};

    const rangeA = m[2];
    const rangeB = m[3];
    const single = m[4];

    const emit = (parts, isRange) => {
      const mapped = parts.map((p, idx) => {
        const np = Object.assign({}, p);
        const newRow = ctx.mapEndpoint(
          np,
          isRange && idx === 1,
          isRange ? parts[0].row : undefined,
          refSheet.name
        );
        if (newRow === null) return null; // other sheet: untouched
        np.row = newRow === undefined ? undefined : newRow;
        return np;
      });
      if (mapped.some(x => x === null)) return;
      if (isRange && mapped[1].row !== undefined && mapped[1].row < mapped[0].row) {
        failures.push(`contains the range "${token}" removed by the edit`);
        return;
      }
      const newPrefix = prefix ? `${prefix}!` : '';
      const body = mapped.map(encodeEndpoint).join(isRange ? ':' : '');
      replacements.push({
        start: matchStart,
        end: matchStart + token.length,
        text: newPrefix + body,
      });
    };

    if (single !== undefined) {
      if (!isValidEndpoint(single)) return;
      const p = parseEndpoint(single);
      if (p.row === undefined) return; // whole-column token on its own
      emit([p], false);
    } else {
      if (!isValidEndpoint(rangeA) || !isValidEndpoint(rangeB)) return;
      const kindA = endpointKind(rangeA);
      if (kindA !== endpointKind(rangeB)) return;
      if (kindA === 'col') return; // whole columns: row edit irrelevant
      emit([parseEndpoint(rangeA), parseEndpoint(rangeB)], true);
    }
  }
}

function skipQuoted(formula, i) {
  let j = i + 1;
  while (j < formula.length) {
    if (formula[j] === '\'') {
      if (formula[j + 1] === '\'') {
        j += 2;
      } else {
        return j + 1;
      }
    } else {
      j++;
    }
  }
  return j;
}

function threeDTouches(span, edit, sheetNames) {
  if (!sheetNames) return true;
  const idx = name => sheetNames.findIndex(s => s.toLowerCase() === name.toLowerCase());
  const firstIdx = idx(span.first);
  const lastIdx = idx(span.last);
  const editIdx = idx(edit.sheetName);
  if (firstIdx === -1 || lastIdx === -1 || editIdx === -1) return true;
  return editIdx >= Math.min(firstIdx, lastIdx) && editIdx <= Math.max(firstIdx, lastIdx);
}

/**
 * Rewrite a sqref string (space-separated A1 ranges without sheet
 * prefix) attached to the edited sheet.
 */
function transformSqref(sqref, edit) {
  return sqref
    .split(/\s+/)
    .filter(Boolean)
    .map(area => {
      const parts = area.split(':');
      if (parts.length === 1) {
        const p = parseSqrefEndpoint(parts[0]);
        p.row = mapIdentityRow(p.row, edit.start, edit.count);
        return p.row === undefined ? null : encodeEndpoint(p);
      }
      const a = parseSqrefEndpoint(parts[0]);
      const b = parseSqrefEndpoint(parts[1]);
      a.row = mapRangeEndpoint('top', a.row, edit.start, edit.count);
      b.row = mapRangeEndpoint('bottom', b.row, edit.start, edit.count, a.row);
      if (b.row < a.row) return null;
      return `${encodeEndpoint(a)}:${encodeEndpoint(b)}`;
    })
    .filter(Boolean)
    .join(' ');
}

function parseSqrefEndpoint(ep) {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})$/.exec(ep);
  if (!m) {
    throw new Error(`Cannot structurally edit workbook: invalid address "${ep}"`);
  }
  return {
    colDollar: m[1] || '',
    col: m[2].toUpperCase(),
    rowDollar: m[3] || '',
    row: parseInt(m[4], 10),
  };
}

module.exports = {
  assertTransformable,
  transformFormula,
  transformRuleFormula,
  transformSqref,
  mapIdentityRow,
};

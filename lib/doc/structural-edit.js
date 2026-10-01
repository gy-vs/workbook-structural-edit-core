'use strict';

const {
  assertTransformable,
  transformFormula,
  transformRuleFormula,
  transformSqref,
  mapIdentityRow,
} = require('../utils/structure-transform');

// Workbook-level, structural row insertion.
//
// Unlike Worksheet.spliceRows/insertRows (which only move cell contents
// on one sheet), a structural insert treats the rows on and after the
// insert point as *identities* moving to new row numbers, and keeps the
// whole workbook consistent with that move:
//
//   - formulas on every sheet that reference the edited cells/ranges
//     (absolute row markers included - the referenced cells physically
//     moved), including same-sheet and range references; references to
//     other sheets and text inside quoted strings are left alone
//   - defined names (the physical splice already maintains these, so it
//     is invoked in a way that moves them exactly once)
//   - data validations (sqref and rule formulae)
//   - conditional formatting (ref and rule formulae)
//   - shared/array formula metadata, auto filters, tables, page breaks
//
// Everything that could fail is validated before any mutation: a
// reference that cannot be maintained reliably (INDIRECT, external
// workbook links, 3D references spanning the edited sheet) aborts the
// whole operation with an explicit error instead of producing a
// valid-looking workbook pointing at wrong data.

const ROW_RX = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})$/;

function parseA1(ref) {
  const m = ROW_RX.exec(ref);
  if (!m) {
    throw new Error(`Cannot structurally edit workbook: invalid address "${ref}"`);
  }
  return {
    colDollar: m[1] || '',
    col: m[2].toUpperCase(),
    rowDollar: m[3] || '',
    row: parseInt(m[4], 10),
  };
}

function buildA1(p) {
  return `${p.colDollar || ''}${p.col}${p.rowDollar || ''}${p.row}`;
}

// Bounding rows of a sqref (space-separated areas).
function sqrefRows(sqref) {
  let top = Infinity;
  let bottom = -Infinity;
  sqref
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .forEach(area => {
      const parts = area.split(':');
      const first = parseA1(parts[0]);
      const last = parseA1(parts[parts.length - 1]);
      top = Math.min(top, first.row);
      bottom = Math.max(bottom, last.row);
    });
  return {top, bottom};
}

function structuralSpliceRows(workbook, worksheet, start, count, inserts, options) {
  if (!worksheet) {
    throw new Error('Cannot structurally edit workbook: worksheet not found');
  }
  if (!Number.isInteger(start) || start < 1) {
    throw new Error(
      `Cannot structurally edit "${worksheet.name}": start must be a 1-based row number`
    );
  }
  if (!Number.isInteger(count) || count < 0) {
    throw new Error(
      `Cannot structurally edit "${worksheet.name}": count must be a non-negative integer`
    );
  }
  if (count === 0) {
    return [];
  }

  const edit = {sheetName: worksheet.name, start, count};
  const sheetNames = workbook.worksheets.map(ws => ws.name);
  const inherit =
    options && options.inheritValidations !== undefined
      ? Boolean(options.inheritValidations)
      : true;
  const inheritStyles =
    options && options.inheritStyles !== undefined ? Boolean(options.inheritStyles) : false;

  // ---------------------------------------------------------------
  // Phase 1: validate everything that could fail, before mutating.
  // ---------------------------------------------------------------
  workbook.worksheets.forEach(ws => {
    ws.eachRow({includeEmpty: false}, row => {
      row.eachCell({includeEmpty: false}, cell => {
        const {value} = cell;
        if (value && typeof value === 'object' && value.formula) {
          assertTransformable(value.formula, ws.name);
          transformFormula(value.formula, edit, ws.name, sheetNames);
        }
      });
    });
    // rule formulae on every sheet (they may reference the edited sheet)
    validateRuleFormulae(collectDataValidationFormulae(ws), ws, edit, sheetNames);
    validateRuleFormulae(collectConditionalFormattingFormulae(ws), ws, edit, sheetNames);
  });

  // Defined-name ranges that cross the insert point. The existing
  // definedNames.spliceRows shifts individual cells (splitting such a
  // range); we remember the crossing ranges so we can fill the newly
  // inserted gap afterwards and keep expressing one contiguous region
  // of the original data plus the new rows.
  const crossingNameRanges = collectCrossingNameRanges(workbook, edit);

  // ---------------------------------------------------------------
  // Phase 2: rewrite formulas on the edited sheet, then perform the
  // physical splice (which also adjusts defined names, once).
  // ---------------------------------------------------------------
  rewriteWorksheetFormulas(worksheet, edit, sheetNames);
  const masterMoves = collectSharedMasterMoves(worksheet, start, count);

  // Worksheet._spliceRows only shifts rows when it has row values to put
  // in the new slots - pad with null so a count of blank rows performs a
  // real physical insertion (and adjusts defined names exactly once).
  const padded = [];
  for (let i = 0; i < count; i++) {
    padded.push(i < inserts.length ? inserts[i] : null);
  }
  worksheet._spliceRows(start, 0, padded, true);

  repairFormulaMetadata(worksheet, edit, masterMoves);
  fillDefinedNameGaps(workbook, crossingNameRanges, edit);

  if (inheritStyles) {
    copyStylesFromAbove(worksheet, start, count);
  }

  // ---------------------------------------------------------------
  // Phase 3: formulas on every other sheet.
  // ---------------------------------------------------------------
  workbook.worksheets.forEach(ws => {
    if (ws !== worksheet) {
      rewriteWorksheetFormulas(ws, edit, sheetNames);
    }
  });

  // ---------------------------------------------------------------
  // Phase 4: structural metadata.
  // ---------------------------------------------------------------
  workbook.worksheets.forEach(ws => {
    shiftRowBreaks(ws, edit);
    if (ws === worksheet) {
      adjustAutoFilter(ws, edit);
      adjustTables(ws, edit);
    }
    // sqref moves on the edited sheet; rule formulae move on every sheet
    rewriteRuleFormulae(ws, edit, sheetNames);
    if (ws === worksheet) {
      adjustDataValidations(ws, edit, inherit);
      adjustConditionalFormattings(ws, edit, inherit);
    }
  });

  return worksheet.getRows(start, count);
}

// ---------------------------------------------------------------------
// Cell formulas
// ---------------------------------------------------------------------

function rewriteWorksheetFormulas(worksheet, edit, sheetNames) {
  worksheet.eachRow({includeEmpty: false}, row => {
    row.eachCell({includeEmpty: false}, cell => {
      const {value} = cell;
      if (value && typeof value === 'object' && value.formula) {
        const rewritten = transformFormula(value.formula, edit, worksheet.name, sheetNames);
        if (rewritten !== value.formula) {
          cell.value = Object.assign({}, value, {formula: rewritten});
        }
      }
    });
  });
}

// Shared-formula masters carry an address + ref range. Record where each
// master moves so slave pointers can be repaired after the splice.
function collectSharedMasterMoves(worksheet, start, count) {
  const moves = {};
  worksheet.eachRow({includeEmpty: false}, row => {
    row.eachCell({includeEmpty: false}, cell => {
      const {value} = cell;
      if (
        value &&
        typeof value === 'object' &&
        value.shareType === 'shared' &&
        value.ref
      ) {
        const newRow = mapIdentityRow(row.number, start, count);
        if (newRow !== undefined) {
          moves[cell.address] = cell.address.replace(/[0-9]+$/, String(newRow));
        }
      }
    });
  });
  return moves;
}

function repairFormulaMetadata(worksheet, edit, masterMoves) {
  worksheet.eachRow({includeEmpty: false}, row => {
    row.eachCell({includeEmpty: false}, cell => {
      const {value} = cell;
      if (!value || typeof value !== 'object') return;
      let next = null;
      if (value.sharedFormula && Object.prototype.hasOwnProperty.call(masterMoves, value.sharedFormula)) {
        next = Object.assign(next || {}, value, {sharedFormula: masterMoves[value.sharedFormula]});
      }
      if (
        value.ref &&
        (value.shareType === 'shared' || value.shareType === 'array')
      ) {
        next = Object.assign(next || {}, value, {ref: transformSqref(value.ref, edit)});
      }
      if (next) {
        cell.value = next;
      }
    });
  });
}

// ---------------------------------------------------------------------
// Validation / conditional-formatting formulae
// ---------------------------------------------------------------------

function collectDataValidationFormulae(worksheet) {
  const formulae = [];
  const seen = new Set();
  Object.keys(worksheet.dataValidations.model || {}).forEach(address => {
    const validation = worksheet.dataValidations.model[address];
    if (!validation || seen.has(validation)) return;
    seen.add(validation);
    (validation.formulae || []).forEach(f => {
      if (typeof f === 'string') formulae.push(f);
    });
  });
  return formulae;
}

function collectConditionalFormattingFormulae(worksheet) {
  const formulae = [];
  (worksheet.conditionalFormattings || []).forEach(cf => {
    (cf.rules || []).forEach(rule => {
      (rule.formulae || []).forEach(f => formulae.push(f));
    });
  });
  return formulae;
}

function validateRuleFormulae(formulae, worksheet, edit, sheetNames) {
  formulae.forEach(formula => {
    assertTransformable(formula, worksheet.name);
    // identity pass catches unmaintainable 3D references; the precise
    // anchor-relative rewrite is applied later.
    transformFormula(formula, edit, worksheet.name, sheetNames);
  });
}

// Rewrite rule template formulae on every worksheet. Rules on the edited
// sheet slide relative references by their anchor's movement; rules on
// other sheets only have cross-sheet references to the edited cells
// rewritten (always as identities).
function rewriteRuleFormulae(worksheet, edit, sheetNames) {
  const isEdited = worksheet.name === edit.sheetName;
  const dvModel = worksheet.dataValidations.model || {};

  // one validation object may back several address keys; find the
  // top-most row for each so the template anchor delta is correct
  const validationTop = new Map();
  Object.keys(dvModel).forEach(address => {
    const validation = dvModel[address];
    if (!validation) return;
    const {top} = sqrefRows(address);
    if (!validationTop.has(validation) || top < validationTop.get(validation)) {
      validationTop.set(validation, top);
    }
  });

  validationTop.forEach((top, validation) => {
    if (!validation.formulae) return;
    const anchorDelta = isEdited && top >= edit.start ? edit.count : 0;
    let changed = false;
    const formulae = validation.formulae.map(formula => {
      if (typeof formula !== 'string') return formula;
      const rewritten = transformRuleFormula(
        formula,
        edit,
        worksheet.name,
        sheetNames,
        anchorDelta
      );
      if (rewritten !== formula) changed = true;
      return rewritten;
    });
    if (changed) {
      validation.formulae = formulae;
    }
  });

  (worksheet.conditionalFormattings || []).forEach(cf => {
    const anchorDelta =
      isEdited && sqrefRows(cf.ref).top >= edit.start ? edit.count : 0;
    (cf.rules || []).forEach(rule => {
      if (!rule.formulae) return;
      rule.formulae = rule.formulae.map(formula =>
        transformRuleFormula(formula, edit, worksheet.name, sheetNames, anchorDelta)
      );
    });
  });
}

// ---------------------------------------------------------------------
// Defined names
// ---------------------------------------------------------------------

// Defined-name cell/range ranges on the edited sheet that span the insert
// point (top < start <= bottom). After definedNames.spliceRows shifts the
// individual cells, such a range is split; fillDefinedNameGaps restores
// its contiguity by marking the new rows as belonging to the name too.
function collectCrossingNameRanges(workbook, edit) {
  const result = [];
  workbook.definedNames.model.forEach(({name, ranges}) => {
    ranges.forEach(rangeStr => {
      const decoded = decodeRange(rangeStr);
      if (!decoded || decoded.sheetName !== edit.sheetName) return;
      if (decoded.top < edit.start && decoded.bottom >= edit.start) {
        result.push({
          name,
          left: decoded.left,
          right: decoded.right,
          top: decoded.top,
          bottom: decoded.bottom,
        });
      }
    });
  });
  return result;
}

function fillDefinedNameGaps(workbook, crossingRanges, edit) {
  const {definedNames} = workbook;
  crossingRanges.forEach(({name, left, right}) => {
    for (let col = left; col <= right; col++) {
      for (let row = edit.start; row < edit.start + edit.count; row++) {
        const colLetters = columnNumberToName(col);
        definedNames.add(`${quoteName(edit.sheetName)}!$${colLetters}$${row}`, name);
      }
    }
  });
}

function decodeRange(rangeStr) {
  // expects Sheet!$A$1:$B$9 or Sheet!$A$1
  const bang = rangeStr.lastIndexOf('!');
  if (bang < 0) return undefined;
  const sheetName = unquoteName(rangeStr.substring(0, bang));
  const refs = rangeStr.substring(bang + 1).replace(/\$/g, '').split(':');
  let a;
  let b;
  try {
    a = parseA1(refs[0]);
    if (refs.length > 1) {
      b = parseA1(refs[1]);
    }
  } catch (err) {
    // whole-column (or otherwise non-cell) defined ranges are shifted by
    // the existing definedNames splice where applicable; nothing to fill
    return undefined;
  }
  if (refs.length === 1) {
    return {
      sheetName,
      left: columnNameToNumber(a.col),
      right: columnNameToNumber(a.col),
      top: a.row,
      bottom: a.row,
    };
  }
  return {
    sheetName,
    left: columnNameToNumber(a.col),
    right: columnNameToNumber(b.col),
    top: a.row,
    bottom: b.row,
  };
}

function quoteName(name) {
  return /^[A-Za-z0-9_.]+$/.test(name) ? name : `'${name.replace(/'/g, '\'\'')}'`;
}

function unquoteName(name) {
  return name.startsWith('\'') && name.endsWith('\'')
    ? name.slice(1, -1).replace(/''/g, '\'')
    : name;
}

function columnNameToNumber(letters) {
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    n = (n * 26) + (letters.toUpperCase().charCodeAt(i) - 64);
  }
  return n;
}

function columnNumberToName(number) {
  let n = number;
  let letters = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

// ---------------------------------------------------------------------
// Data validations
// ---------------------------------------------------------------------

// Inheritance policy (explicit and consistent):
//   - rule ranges moved with their cells when the insert is before them
//   - rule ranges crossed by an insertion expand to include the new rows
//   - single-cell rules move with their cell and never spread
// New rows inserted inside (or at the top edge of) a range-based rule
// therefore inherit it.
function adjustDataValidations(worksheet, edit, sheetNames, inherit) {
  const oldModel = worksheet.dataValidations.model || {};
  const addresses = Object.keys(oldModel).filter(a => oldModel[a]);

  // group addresses by the shared validation object (the xform expands
  // one sqref into many keys pointing at one instance)
  const groups = new Map();
  addresses.forEach(address => {
    const validation = oldModel[address];
    if (!groups.has(validation)) {
      groups.set(validation, []);
    }
    groups.get(validation).push(address);
  });

  const newModel = {};
  groups.forEach((keys, validation) => {
    let top = Infinity;
    let bottom = -Infinity;
    keys.forEach(address => {
      const {top: t, bottom: b} = sqrefRows(address);
      top = Math.min(top, t);
      bottom = Math.max(bottom, b);
    });

    const anchorMoves = top >= edit.start;
    const crossed = !anchorMoves && bottom >= edit.start;

    keys.forEach(address => {
      if (address.includes(':')) {
        // crossing a range expands it; if inheritance is disabled, keep
        // the original extent (bottom does not include the new rows)
        const target = crossed && !inherit ? address : transformSqref(address, edit);
        if (target) {
          target
            .trim()
            .split(/\s+/)
            .filter(Boolean)
            .forEach(area => assignRange(newModel, area, validation));
        }
      } else {
        const p = parseA1(address);
        const newRow = mapIdentityRow(p.row, edit.start, edit.count);
        if (newRow !== undefined) {
          p.row = newRow;
          newModel[buildA1(p)] = validation;
        }
      }
    });
  });

  worksheet.dataValidations.model = newModel;
}

function assignRange(model, area, validation) {
  // the model is keyed per cell by the xform conventions
  const parts = area.split(':');
  if (parts.length === 1) {
    model[parts[0]] = validation;
    return;
  }
  const a = parseA1(parts[0]);
  const b = parseA1(parts[1]);
  const colToNumber = letters => {
    let n = 0;
    for (let i = 0; i < letters.length; i++) {
      n = (n * 26) + (letters.charCodeAt(i) - 64);
    }
    return n;
  };
  const nToCol = number => {
    let n = number;
    let letters = '';
    while (n > 0) {
      const rem = (n - 1) % 26;
      letters = String.fromCharCode(65 + rem) + letters;
      n = Math.floor((n - 1) / 26);
    }
    return letters;
  };
  for (let c = colToNumber(a.col); c <= colToNumber(b.col); c++) {
    for (let r = a.row; r <= b.row; r++) {
      model[`${nToCol(c)}${r}`] = validation;
    }
  }
}

// ---------------------------------------------------------------------
// Conditional formatting
// ---------------------------------------------------------------------

// Rule formulae are rewritten separately by rewriteRuleFormulae (which
// runs on every sheet); here only the sqref moves with the business data.
function adjustConditionalFormattings(worksheet, edit, inherit) {
  worksheet.conditionalFormattings = (worksheet.conditionalFormattings || [])
    .map(cf => {
      const {top, bottom} = sqrefRows(cf.ref);
      const anchorMoves = top >= edit.start;
      const crossed = !anchorMoves && bottom >= edit.start;

      // keep original business extent when crossing without inheritance
      const ref = crossed && !inherit ? cf.ref : transformSqref(cf.ref, edit);
      return ref ? Object.assign({}, cf, {ref}) : null;
    })
    .filter(Boolean);
}

// ---------------------------------------------------------------------
// Styles, auto filter, tables, page breaks
// ---------------------------------------------------------------------

function copyStylesFromAbove(worksheet, start, count) {
  for (let i = 0; i < count; i++) {
    const above = worksheet.findRow(start + i - 1);
    const target = worksheet.getRow(start + i);
    if (!above) continue;
    if (above.height) {
      target.height = above.height;
    }
    target.style = JSON.parse(JSON.stringify(above.style || {}));
    above.eachCell({includeEmpty: true}, (cell, colNumber) => {
      target.getCell(colNumber).style = JSON.parse(JSON.stringify(cell.style || {}));
    });
  }
}

function adjustAutoFilter(worksheet, edit) {
  const {autoFilter} = worksheet;
  if (!autoFilter) return;
  if (typeof autoFilter === 'string') {
    worksheet.autoFilter = transformSqref(autoFilter, edit);
  } else if (autoFilter.from && autoFilter.to) {
    const shiftAddress = address => {
      const p = parseA1(address);
      p.row = p.row >= edit.start ? p.row + edit.count : p.row;
      return buildA1(p);
    };
    autoFilter.from = shiftAddress(autoFilter.from);
    autoFilter.to = shiftAddress(autoFilter.to);
  }
}

function adjustTables(worksheet, edit) {
  worksheet.getTables().forEach(table => {
    if (!table.ref || !table.ref.includes(':')) return;
    const parts = table.ref.split(':');
    const a = parseA1(parts[0]);
    const b = parseA1(parts[1]);
    a.row = a.row >= edit.start ? a.row + edit.count : a.row;
    b.row = b.row >= edit.start ? b.row + edit.count : b.row;
    table.ref = `${buildA1(a)}:${buildA1(b)}`;
  });
}

function shiftRowBreaks(worksheet, edit) {
  if (!worksheet.rowBreaks) return;
  worksheet.rowBreaks.forEach(brk => {
    if (typeof brk.id === 'number' && brk.id >= edit.start) {
      brk.id += edit.count;
    }
  });
}

module.exports = {
  structuralSpliceRows,
};

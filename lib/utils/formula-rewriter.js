'use strict';

// Structural editing support for formulas and address-like expressions.
//
// When rows are inserted into one worksheet, every reference anywhere in the
// workbook that points at cells of that sheet must be updated so it still
// points at the same business data. This module performs that rewrite without
// a full formula parser: it scans a formula as text, skipping quoted strings
// and recognising sheet-qualified references, and shifts only the references
// that belong to the sheet being edited.
//
// References to any other sheet are copied verbatim, as is text inside
// double-quoted strings. References that cannot be maintained reliably while
// pointing at the edited sheet (3D spans covering the sheet, R1C1 style)
// raise an explicit error instead of silently rewriting to the wrong cell.

// sticky scanner alternatives (all anchored at the current index):
//  token 0: double-quoted string literal
//  token 1: external reference (e.g. [1]Sheet1!A1)
//  group 1-3: 3D reference spanning multiple sheets (e.g. Sheet1:Sheet2!A1)
//  group 4-5: single-sheet prefix, quoted or unquoted ('My Sheet!' or Sheet1!)
const tokenRx = new RegExp(
  [
    '"(?:""|[^"])*"',
    '\\[(?:\\d+|[^\\]!]*)\\][^!]*!',
    '([A-Za-z0-9_.\\-]+)\\s*:\\s*(?:\'((?:[^\']|\'\')*)\'|([A-Za-z0-9_.\\- ]+))!',
    '(?:\'((?:[^\']|\'\')*)\'|([A-Za-z0-9_.\\-]+))!',
  ].join('|'),
  'gy'
);

// local A1 reference immediately following a matched sheet prefix.
// whole rows (1:10) and whole columns (A:B) are accepted too.
const prefixedRefRx =
  /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)?(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+)?)?/;

// bare local A1 cell or range reference (must name a row)
const localRefRx = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?/;

// bare whole-row span - a colon is mandatory so plain numeric literals and
// row numbers (100) are not mistaken for references
const wholeRowRx = /^(\$?)(\d+):(\$?)(\d+)(?![0-9])/;

// a token starting with R...C that looks like R1C1 style
const r1c1StartRx = /^\$?[RC]/i;
const r1c1Rx = /^\$?R(?:\[-?\d+]|\d+)?\$?C(?:\[-?\d+]|\d+)?/i;

// character preceding a bare reference must not be an identifier part,
// otherwise the token is really a name (e.g. the X1 in FOOX1) or a number
const notIdentPartRx = /[^A-Za-z0-9_.]/;

function sameSheetName(a, b) {
  return a !== undefined && a.toUpperCase() === b.toUpperCase();
}

function shiftRow(row, insertAt, count) {
  return row >= insertAt ? row + count : row;
}

function shiftRef(match, insertAt, count) {
  // groups: 1 dCol1, 2 col1, 3 dRow1, 4 row1,
  //         5 dCol2, 6 col2, 7 dRow2, 8 row2
  const row1 = match[4] === undefined ? undefined : parseInt(match[4], 10);
  const row2 = match[8] === undefined ? undefined : parseInt(match[8], 10);

  // whole-column references name no rows - inserting rows never changes them
  if (row1 === undefined) {
    return match[0];
  }

  let result = match[1] + match[2].toUpperCase() + match[3] + shiftRow(row1, insertAt, count);
  if (match[6] !== undefined) {
    result +=
      `:${
      match[5]
      }${match[6].toUpperCase()
      }${match[7]
      }${row2 === undefined ? '' : shiftRow(row2, insertAt, count)}`;
  }
  return result;
}

function shiftWholeRow(match, insertAt, count) {
  return (
    `${match[1] +
    shiftRow(parseInt(match[2], 10), insertAt, count)
    }:${
    match[3]
    }${shiftRow(parseInt(match[4], 10), insertAt, count)}`
  );
}

function isLikelyFunctionCall(formula, index, length) {
  // a reference used as a function callee is followed by '('
  return formula[index + length] === '(';
}

function rewritePrefixedRef(formula, index, token, prefixSheet, ctx) {
  const {targetSheet, insertAt, count, colCache} = ctx;
  const rest = formula.slice(index);

  const isTarget = sameSheetName(prefixSheet, targetSheet);

  // prefixed whole-row references like Sheet!1:10 - checked before the
  // column-based A1 pattern, which cannot match a leading digit
  const wholeRowMatch = rest.match(wholeRowRx);
  if (wholeRowMatch) {
    return isTarget
      ? {length: wholeRowMatch[0].length, text: token + shiftWholeRow(wholeRowMatch, insertAt, count)}
      : {length: wholeRowMatch[0].length, text: token + wholeRowMatch[0]};
  }

  const match = rest.match(prefixedRefRx);

  if (!match) {
    // no address-like token after the prefix - leave the prefix in place
    return null;
  }

  // any sheet-qualified A1 reference: validate column tokens regardless of
  // which sheet it targets
  if (r1c1Rx.test(rest)) {
    if (isTarget) {
      throw new Error(
        `Cannot maintain R1C1-style reference in formula while inserting rows into '${targetSheet}': ${formula}`
      );
    }
    return {length: match[0].length, text: token + match[0]};
  }

  [match[2], match[6]].forEach(col => {
    if (col !== undefined) {
      colCache.l2n(col.toUpperCase()); // validates 1..XFD
    }
  });

  if (isLikelyFunctionCall(formula, index, match[0].length)) {
    // e.g. Sheet!NAME( - consume the look-alike so the bare scanner never
    // touches its row number
    return {length: match[0].length, text: token + match[0]};
  }

  return isTarget
    ? {length: match[0].length, text: token + shiftRef(match, insertAt, count)}
    : {length: match[0].length, text: token + match[0]};
}

function rewriteLocalRef(formula, index, ctx) {
  const {insertAt, count, targetSheet} = ctx;
  const rest = formula.slice(index);
  const prev = formula[index - 1];
  const guardOk = prev === undefined || notIdentPartRx.test(prev);
  const startsDigit = /^(\$?)\d/.test(rest);

  // references starting with a digit can only be whole-row spans (1:10)
  if (startsDigit) {
    const prevDigit = prev !== undefined && /[0-9.]/.test(prev);
    const match = rest.match(wholeRowRx);
    if (match && !prevDigit && !isLikelyFunctionCall(formula, index, match[0].length)) {
      return {length: match[0].length, text: shiftWholeRow(match, insertAt, count)};
    }
    return null;
  }

  // R1C1-style references only ever address the owning sheet
  if (r1c1StartRx.test(rest)) {
    const r1c1Match = rest.match(r1c1Rx);
    if (r1c1Match && guardOk && !isLikelyFunctionCall(formula, index, r1c1Match[0].length)) {
      throw new Error(
        `Cannot maintain R1C1-style reference in formula while inserting rows into '${targetSheet}': ${formula}`
      );
    }
  }

  const match = rest.match(localRefRx);
  if (match && guardOk && !isLikelyFunctionCall(formula, index, match[0].length)) {
    return {length: match[0].length, text: shiftRef(match, insertAt, count)};
  }

  return null;
}

// shift every reference to targetSheet inside a single formula string.
// options:
//   ownSheet:        the formula lives on targetSheet, so bare references
//                    implicitly address it (otherwise they are left alone)
//   is3DCovering:    predicate(firstSheet, secondSheet) -> true when a 3D
//                    span includes targetSheet (only such spans can fail)
function adjustFormula(formula, targetSheet, insertAt, count, colCache, options = {}) {
  if (typeof formula !== 'string' || !formula) {
    return formula;
  }
  const ctx = {targetSheet, insertAt, count, colCache};
  const ownSheet = options.ownSheet !== false;
  const is3DCovering = options.is3DCovering || (() => false);

  let result = '';
  let i = 0;
  tokenRx.lastIndex = 0;

  while (i < formula.length) {
    tokenRx.lastIndex = i;
    const tokenMatch = tokenRx.exec(formula);

    if (tokenMatch && tokenMatch.index === i) {
      const token = tokenMatch[0];
      const after = i + token.length;

      // double-quoted string literal: contains no references
      if (token.charAt(0) === '"') {
        result += token;
        i = after;
        continue;
      }

      // external reference ([1]... or [Book.xlsx]...): never local
      if (token.charAt(0) === '[') {
        result += token;
        i = after;
        continue;
      }

      // 3D sheet span - only fail when the span actually covers the edit
      if (tokenMatch[1] !== undefined) {
        const first = tokenMatch[1];
        const second = tokenMatch[2] !== undefined ? tokenMatch[2].replace(/''/g, '\'') : tokenMatch[3];
        if (is3DCovering(first, second)) {
          throw new Error(
            `Cannot maintain 3D reference '${token}' while inserting rows into '${targetSheet}': ${formula}`
          );
        }
        result += token;
        i = after;
        continue;
      }

      // single sheet prefix - inspect the reference that follows it
      const prefixSheet =
        tokenMatch[4] !== undefined ? tokenMatch[4].replace(/''/g, '\'') : tokenMatch[5];
      const rewrite = rewritePrefixedRef(formula, after, token, prefixSheet, ctx);
      if (rewrite) {
        result += rewrite.text;
        i = after + rewrite.length;
      } else {
        result += token;
        i = after;
      }
      continue;
    }

    // bare local references only belong to the sheet that owns the formula
    if (ownSheet) {
      const rewrite = rewriteLocalRef(formula, i, ctx);
      if (rewrite) {
        result += rewrite.text;
        i += rewrite.length;
        continue;
      }
    }

    result += formula[i];
    i += 1;
  }

  return result;
}

module.exports = {
  adjustFormula,
};

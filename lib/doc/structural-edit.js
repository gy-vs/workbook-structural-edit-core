'use strict';

// Workbook-level structural row insertion.
//
// Unlike Worksheet#insertRows/spliceRows (which only move cell data inside
// one sheet), this operation treats the row insert as a change to workbook
// structure. References that pointed at content below the insert position are
// rewritten so they keep pointing at the same cells:
//
//   * formulas in every worksheet (own-sheet and cross-sheet references,
//     absolute row markers included) - only references to the edited sheet
//     are touched, text inside double-quoted strings is never a reference;
//   * shared/array formula anchors;
//   * defined names that resolve to cell ranges (existing spliceRows
//     bookkeeping is bypassed so names are moved exactly once);
//   * data validation sqref areas and their formulae;
//   * conditional formatting sqref areas and rule formulae;
//   * merged cell ranges.
//
// Each range endpoint shifts independently (Excel semantics):
//   insert before the range   -> range shifts down;
//   insert at the first row    -> range shifts down (identity moves);
//   insert strictly inside     -> range grows by the inserted rows;
//   insert at/after the end    -> unchanged.
//
// References that cannot be maintained reliably (3D spans covering the
// edited sheet, R1C1 references into it) throw rather than produce a workbook
// that looks valid but points elsewhere.

const Enums = require('./enums');
const colCache = require('../utils/col-cache');
const {adjustFormula} = require('../utils/formula-rewriter');

class StructuralEditError extends Error {}

// ---------------------------------------------------------------------------
// range shifting primitives
// ---------------------------------------------------------------------------

function shiftRow(row, insertAt, count) {
  return row >= insertAt ? row + count : row;
}

// shift one bare A1 address (possibly with $ markers), same sheet
function shiftAddress(addressStr, insertAt, count) {
  const match = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(addressStr);
  if (!match) {
    return addressStr;
  }
  return (
    match[1] + match[2].toUpperCase() + match[3] + shiftRow(parseInt(match[4], 10), insertAt, count)
  );
}

// shift one A1 range with no sheet prefix (e.g. "$A$2:$A$3" or "B3")
function shiftLocalRangeText(text, insertAt, count) {
  return text
    .split(':')
    .map(part => shiftAddress(part.trim(), insertAt, count))
    .join(':');
}

// shift the pieces of a whitespace-separated sqref. Pieces with an explicit
// sheet prefix other than targetSheet are left alone; bare pieces are only
// shifted when localIsTarget is true.
function shiftSqref(sqref, insertAt, count, targetSheet, localIsTarget) {
  return sqref
    .split(/\s+/)
    .filter(Boolean)
    .map(piece => {
      const bang = piece.indexOf('!');
      if (bang >= 0) {
        const sheetPart = piece.slice(0, bang);
        const sheetName = sheetPart.charAt(0) === '\''
          ? sheetPart.slice(1, -1).replace(/''/g, '\'')
          : sheetPart;
        if (sheetName.toUpperCase() !== targetSheet.toUpperCase()) {
          return piece;
        }
      } else if (!localIsTarget) {
        return piece;
      }
      const refPart = bang >= 0 ? piece.slice(bang + 1) : piece;
      const prefix = bang >= 0 ? piece.slice(0, bang + 1) : '';
      return prefix + shiftLocalRangeText(refPart, insertAt, count);
    })
    .join(' ');
}

// ---------------------------------------------------------------------------
// defined names
// ---------------------------------------------------------------------------

function shiftDefinedNames(definedNames, targetSheet, insertAt, count) {
  const {model} = definedNames; // [{name, ranges}]
  if (!model.length) {
    return;
  }

  definedNames.matrixMap = {};
  model.forEach(entry => {
    entry.ranges.forEach(rangeStr => {
      definedNames.add(
        shiftDefinedNameRange(rangeStr, targetSheet, insertAt, count),
        entry.name
      );
    });
  });
}

function shiftDefinedNameRange(rangeStr, targetSheet, insertAt, count) {
  const decoded = colCache.decodeEx(rangeStr);
  if (!decoded.sheetName) {
    // constants / formula-defined names carry no sheet prefix - leave alone
    return rangeStr;
  }
  if (decoded.sheetName.toUpperCase() !== targetSheet.toUpperCase()) {
    return rangeStr;
  }
  if (decoded.error) {
    return rangeStr;
  }
  const bang = rangeStr.indexOf('!');
  return rangeStr.slice(0, bang + 1) + shiftLocalRangeText(rangeStr.slice(bang + 1), insertAt, count);
}

// ---------------------------------------------------------------------------
// formulas
// ---------------------------------------------------------------------------

function adjustAllFormulas(workbook, targetSheet, insertAt, count) {
  const names = workbook.worksheets.map(ws => ws.name.toUpperCase());
  const target = targetSheet.toUpperCase();

  const covers3D = (first, second) => {
    const ia = names.indexOf(first.toUpperCase());
    const ib = names.indexOf(second.toUpperCase());
    if (ia === -1 || ib === -1) {
      return false;
    }
    const lo = Math.min(ia, ib);
    const hi = Math.max(ia, ib);
    const it = names.indexOf(target);
    return it >= lo && it <= hi;
  };

  workbook.worksheets.forEach(worksheet => {
    const ownSheet = worksheet.name.toUpperCase() === target;

    const rewrite = formula =>
      adjustFormula(formula, targetSheet, insertAt, count, colCache, {
        ownSheet,
        is3DCovering: covers3D,
      });

    worksheet.eachRow({includeEmpty: true}, row => {
      row.eachCell({includeEmpty: true}, cell => {
        if (cell.type !== Enums.ValueType.Formula) {
          return;
        }
        const valueModel = cell.value;
        let changed = false;

        if (valueModel.formula) {
          valueModel.formula = rewrite(valueModel.formula);
          changed = true;
        }
        if (valueModel.sharedFormula) {
          valueModel.sharedFormula = shiftSqref(
            valueModel.sharedFormula,
            insertAt,
            count,
            targetSheet,
            ownSheet
          ).replace(/\s+/g, '');
          changed = true;
        }
        if (valueModel.ref) {
          valueModel.ref = shiftSqref(valueModel.ref, insertAt, count, targetSheet, ownSheet);
          changed = true;
        }

        if (changed) {
          cell.value = valueModel;
        }
      });
    });
  });
}

// ---------------------------------------------------------------------------
// data validations
// ---------------------------------------------------------------------------

function adjustDataValidations(worksheet, targetSheet, insertAt, count, inheritDataValidations) {
  const oldModel = worksheet.dataValidations.model;
  const newModel = {};

  const place = (address, validation) => {
    newModel[shiftLocalRangeText(address, insertAt, count)] = validation;
  };

  Object.keys(oldModel).forEach(address => {
    const validation = oldModel[address];
    if (!validation) {
      return;
    }

    let shiftedValidation = validation;
    if (validation.formulae) {
      shiftedValidation = Object.assign({}, validation, {
        formulae: validation.formulae.map(f =>
          typeof f === 'string'
            ? adjustFormula(f, targetSheet, insertAt, count, colCache, {ownSheet: true})
            : f
        ),
      });
    }

    // normalise to individual addresses, matching the representation produced
    // when an xlsx file is read, so in-memory and reloaded results agree
    const decoded = colCache.decodeEx(address);
    if (decoded.top) {
      for (let r = decoded.top; r <= decoded.bottom; r++) {
        for (let c = decoded.left; c <= decoded.right; c++) {
          place(colCache.encodeAddress(r, c), shiftedValidation);
        }
      }
      // when the insert cuts through the validation area, the new blank rows
      // inherit the same rule (Excel expands the validation region)
      if (inheritDataValidations && decoded.top < insertAt && decoded.bottom >= insertAt) {
        for (let nr = insertAt; nr < insertAt + count; nr++) {
          for (let c = decoded.left; c <= decoded.right; c++) {
            newModel[colCache.encodeAddress(nr, c)] = shiftedValidation;
          }
        }
      }
    } else {
      place(address, shiftedValidation);
    }
  });

  worksheet.dataValidations.model = newModel;
}

// ---------------------------------------------------------------------------
// conditional formatting
// ---------------------------------------------------------------------------

function adjustConditionalFormatting(worksheet, targetSheet, insertAt, count) {
  const rewrite = formula =>
    adjustFormula(formula, targetSheet, insertAt, count, colCache, {ownSheet: true});

  worksheet.conditionalFormattings.forEach(cf => {
    if (cf.ref) {
      cf.ref = shiftSqref(cf.ref, insertAt, count, targetSheet, true);
    }
    (cf.rules || []).forEach(rule => {
      if (rule.formulae) {
        rule.formulae = rule.formulae.map(f => (typeof f === 'string' ? rewrite(f) : f));
      }
      if (rule.ref) {
        rule.ref = shiftSqref(rule.ref, insertAt, count, targetSheet, true);
      }
      if (Array.isArray(rule.cfvo)) {
        rule.cfvo.forEach(cfvo => {
          if (cfvo && cfvo.type === 'formula' && typeof cfvo.value === 'string') {
            cfvo.value = rewrite(cfvo.value);
          }
        });
      }
    });
  });
}

// ---------------------------------------------------------------------------
// merged cells
// ---------------------------------------------------------------------------

function adjustMerges(worksheet, insertAt, count) {
  const oldRanges = Object.values(worksheet._merges).filter(Boolean);
  if (!oldRanges.length) {
    return;
  }
  worksheet._merges = {};

  oldRanges.forEach(range => {
    // both endpoints follow the same rule as every other range
    const shifted = colCache.decode(
      shiftLocalRangeText(`${range.tl}:${range.br}`, insertAt, count)
    );
    worksheet.mergeCells(shifted.top, shifted.left, shifted.bottom, shifted.right);
  });
}

// ---------------------------------------------------------------------------
// tables and auto filter
// ---------------------------------------------------------------------------

function adjustTables(worksheet, insertAt, count) {
  Object.values(worksheet.tables || {}).forEach(table => {
    const model = table.table || table;
    // update every row-bearing footprint the writer uses. do not go through
    // table.ref's setter: it schedules a cell rewrite against the old range.
    ['ref', 'tableRef', 'autoFilterRef'].forEach(key => {
      if (model[key]) {
        model[key] = shiftLocalRangeText(model[key], insertAt, count);
      }
    });
    if (model.tl) {
      model.tl = colCache.decodeAddress(
        shiftAddress(colCache.encodeAddress(model.tl.row, model.tl.col), insertAt, count)
      );
    }
    // a cache created before the edit would rewrite cells at the old footprint
    table._cache = undefined;
  });
}

function shiftModelAddress(address, insertAt, count) {
  if (address === undefined || address === null) {
    return address;
  }
  if (typeof address === 'string') {
    return shiftAddress(address, insertAt, count);
  }
  if (typeof address.row === 'number') {
    return Object.assign({}, address, {row: shiftRow(address.row, insertAt, count)});
  }
  if (typeof address === 'number') {
    return shiftRow(address, insertAt, count);
  }
  return address;
}

function adjustAutoFilter(worksheet, insertAt, count) {
  const {autoFilter} = worksheet;
  if (!autoFilter) {
    return;
  }
  if (typeof autoFilter === 'string') {
    worksheet.autoFilter = shiftLocalRangeText(autoFilter, insertAt, count);
  } else if (autoFilter.from !== undefined && autoFilter.to !== undefined) {
    autoFilter.from = shiftModelAddress(autoFilter.from, insertAt, count);
    autoFilter.to = shiftModelAddress(autoFilter.to, insertAt, count);
  }
}

// ---------------------------------------------------------------------------
// entry point
// ---------------------------------------------------------------------------

// insert `count` rows at 1-based position `start` on worksheetNameOrId.
// values: optional array of row values for the inserted rows.
// options.inheritDataValidations (default true): new rows inside a validation
// area pick up that validation, matching Excel's UI behaviour.
function structuralInsertRows(workbook, worksheetNameOrId, start, count = 1, values = [], options = {}) {
  if (!workbook || typeof workbook.addWorksheet !== 'function') {
    throw new StructuralEditError('structuralInsertRows requires a Workbook');
  }
  const worksheet = workbook.getWorksheet(worksheetNameOrId);
  if (!worksheet) {
    throw new StructuralEditError(`Worksheet not found: ${worksheetNameOrId}`);
  }
  if (!Number.isInteger(start) || start < 1) {
    throw new StructuralEditError(`Invalid insert position: ${start}`);
  }
  if (!Number.isInteger(count) || count < 1) {
    throw new StructuralEditError(`Invalid row count: ${count}`);
  }
  const inheritDataValidations =
    options.inheritDataValidations === undefined ? true : !!options.inheritDataValidations;

  const targetSheet = worksheet.name;

  // pass 1: rewrite all formulas workbook-wide. done before touching rows so a
  // failure leaves the workbook structurally unchanged.
  adjustAllFormulas(workbook, targetSheet, start, count);

  // pass 2: physical row move, with spliceRows' built-in defined-name move
  // suppressed - the name adjustment below owns that bookkeeping
  const rowValues = [];
  for (let i = 0; i < count; i++) {
    rowValues.push(values[i] === undefined ? null : values[i]);
  }
  worksheet._spliceRowsSkipDefinedNames(start, 0, rowValues);

  // pass 3: structural bookkeeping
  shiftDefinedNames(workbook.definedNames, targetSheet, start, count);
  adjustDataValidations(worksheet, targetSheet, start, count, inheritDataValidations);
  adjustConditionalFormatting(worksheet, targetSheet, start, count);
  adjustMerges(worksheet, start, count);
  adjustTables(worksheet, start, count);
  adjustAutoFilter(worksheet, start, count);

  return worksheet.getRows(start, count);
}

module.exports = {
  structuralInsertRows,
  StructuralEditError,
};

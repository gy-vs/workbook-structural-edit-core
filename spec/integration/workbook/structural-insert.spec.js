const ExcelJS = verquire('exceljs');

async function roundTrip(wb) {
  const buffer = await wb.xlsx.writeBuffer();
  const copy = new ExcelJS.Workbook();
  await copy.xlsx.load(buffer);
  return copy;
}

describe('Structural row insertion', () => {
  function buildWorkbook() {
    const wb = new ExcelJS.Workbook();
    const inputs = wb.addWorksheet('Inputs');
    const summary = wb.addWorksheet('Summary');
    const detail = wb.addWorksheet('Detail Two');

    inputs.getCell('A3').value = {formula: '1+2', result: 3};
    inputs.getCell('B3').value = 'marked';
    inputs.getCell('A4').value = 7;

    summary.getCell('B2').value = {formula: 'Inputs!A3+1', result: 4};
    summary.getCell('C2').value = {formula: 'SUM(Inputs!A2:A5)', result: 10};
    summary.getCell('D2').value = {formula: '\'Detail Two\'!A1', result: 0};
    summary.getCell('E2').value = {formula: '"Inputs!A3"', result: 'Inputs!A3'};

    detail.getCell('A1').value = {formula: 'A5', result: 0};

    wb.definedNames.add('Inputs!$A$3', 'ActiveAmount');
    wb.definedNames.add('Inputs!$A$2:$A$4', 'AmountBlock');
    wb.definedNames.add('\'Detail Two\'!$A$1', 'DetailStart');

    inputs.dataValidations.add('B3', {
      type: 'list',
      allowBlank: true,
      formulae: ['$A$2:$A$3'],
    });
    inputs.dataValidations.add('A2:A3', {type: 'whole', operator: 'between', formulae: [1, 100]});

    inputs.addConditionalFormatting({
      ref: 'A2:A3',
      rules: [{type: 'expression', formulae: ['A2>1'], priority: 1}],
    });

    inputs.mergeCells('A5:B6');
    return wb;
  }

  function expectStructure(wb) {
    const inputs = wb.getWorksheet('Inputs');
    const summary = wb.getWorksheet('Summary');
    const detail = wb.getWorksheet('Detail Two');

    // identity of shifted content
    expect(inputs.getCell('A4').value.formula).to.equal('1+2');
    expect(inputs.getCell('B4').value).to.equal('marked');
    expect(inputs.getCell('A5').value).to.equal(7);

    // formulas follow the original cells
    expect(summary.getCell('B2').value.formula).to.equal('Inputs!A4+1');
    expect(summary.getCell('C2').value.formula).to.equal('SUM(Inputs!A2:A6)');

    // formulas unrelated to the edited sheet are untouched
    expect(summary.getCell('D2').value.formula).to.equal('\'Detail Two\'!A1');
    expect(detail.getCell('A1').value.formula).to.equal('A5');

    // text that merely looks like an address is untouched
    expect(summary.getCell('E2').value.formula).to.equal('"Inputs!A3"');

    // defined names
    expect(wb.definedNames.model.find(n => n.name === 'ActiveAmount').ranges).to.deep.equal([
      'Inputs!$A$4',
    ]);
    expect(wb.definedNames.model.find(n => n.name === 'AmountBlock').ranges).to.deep.equal([
      'Inputs!$A$2:$A$5',
    ]);
    expect(wb.definedNames.model.find(n => n.name === 'DetailStart').ranges).to.deep.equal([
      '\'Detail Two\'!$A$1',
    ]);

    // data validations move with the cells and rule formulae are maintained
    const dv = inputs.dataValidations.model;
    expect(dv.B4).to.exist();
    expect(dv.B4.formulae).to.deep.equal(['$A$2:$A$4']);
    expect(dv.A4).to.exist();
    expect(dv.A3).to.exist(); // new row inside the range inherits the rule

    // conditional formatting range and rule formula
    expect(inputs.conditionalFormattings.map(cf => cf.ref)).to.deep.equal(['A2:A4']);
    expect(inputs.conditionalFormattings[0].rules[0].formulae).to.deep.equal(['A2>1']);

    // merges move with content
    expect(Object.values(inputs._merges).map(r => r.range)).to.deep.equal(['A6:B7']);
  }

  it('maintains references in memory via workbook entry', () => {
    const wb = buildWorkbook();
    wb.structuralInsertRows('Inputs', 3, 1);
    expectStructure(wb);
  });

  it('is also reachable on the worksheet', () => {
    const wb = buildWorkbook();
    wb.getWorksheet('Inputs').structuralInsertRows(3);
    expectStructure(wb);
  });

  it('survives writing to xlsx and reloading', async () => {
    const wb = buildWorkbook();
    wb.structuralInsertRows('Inputs', 3, 1);
    const reloaded = await roundTrip(wb);
    expectStructure(reloaded);
  });

  it('inserts supplied values', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A2').value = 2;
    wb.structuralInsertRows('S', 2, 2, [['one'], ['two']]);
    expect(ws.getCell('A2').value).to.equal('one');
    expect(ws.getCell('A3').value).to.equal('two');
    expect(ws.getCell('A4').value).to.equal(2);
  });

  it('shifts ranges wholesale when inserting before them', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A2').value = 2;
    wb.definedNames.add('S!$A$2:$A$4', 'Block');
    ws.addConditionalFormatting({
      ref: 'A2:A4',
      rules: [{type: 'expression', formulae: ['A2>0'], priority: 1}],
    });

    wb.structuralInsertRows('S', 1, 1);

    expect(ws.getCell('A3').value).to.equal(2);
    expect(wb.definedNames.model[0].ranges).to.deep.equal(['S!$A$3:$A$5']);
    expect(ws.conditionalFormattings[0].ref).to.equal('A3:A5');
    expect(ws.conditionalFormattings[0].rules[0].formulae).to.deep.equal(['A3>0']);
  });

  it('does not double-move defined names through spliceRows', () => {
    const wb = buildWorkbook();
    wb.structuralInsertRows('Inputs', 3, 1);
    expect(wb.definedNames.model.find(n => n.name === 'ActiveAmount').ranges).to.deep.equal([
      'Inputs!$A$4',
    ]);
  });

  it('leaves legacy insertRows semantics untouched', () => {
    const wb = buildWorkbook();
    const inputs = wb.getWorksheet('Inputs');
    const summary = wb.getWorksheet('Summary');
    inputs.insertRows(3, [null]);

    // spliceRows moves defined names but not formulas or validations
    expect(summary.getCell('B2').value.formula).to.equal('Inputs!A3+1');
    expect(wb.definedNames.model.find(n => n.name === 'ActiveAmount').ranges).to.deep.equal([
      'Inputs!$A$4',
    ]);
  });

  it('fails explicitly on unmaintainable R1C1 references', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A3').value = {formula: 'R[-1]C', result: 0};
    expect(() => wb.structuralInsertRows('S', 2, 1)).to.throw();
  });

  it('throws when the worksheet does not exist', () => {
    const wb = new ExcelJS.Workbook();
    expect(() => wb.structuralInsertRows('Nope', 1, 1)).to.throw();
  });

  it('optionally does not extend data validations into new rows', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.dataValidations.add('A2:A4', {type: 'whole', operator: 'between', formulae: [1, 9]});
    wb.structuralInsertRows('S', 3, 1, [], {inheritDataValidations: false});

    const dv = ws.dataValidations.model;
    expect(dv.A2).to.exist();
    expect(dv.A3).to.not.exist();
    expect(dv.A4).to.exist();
    expect(dv.A5).to.exist();
  });

  it('maintains table footprints and auto filter ranges', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'Id';
    ws.getCell('A2').value = 1;
    ws.getCell('A3').value = 2;
    ws.addTable({
      name: 'Amounts',
      ref: 'A1:A3',
      headerRow: true,
      columns: [{name: 'Id'}],
      rows: [[1], [2]],
    });
    ws.autoFilter = 'A2:A3';

    wb.structuralInsertRows('S', 3, 1);

    const table = ws.getTable('Amounts');
    expect(table.table.tableRef).to.equal('A1:A4');
    expect(table.table.autoFilterRef).to.equal('A1:A4');
    expect(ws.autoFilter).to.equal('A2:A4');

    const reloaded = await roundTrip(wb);
    const rws = reloaded.getWorksheet('S');
    expect(rws.getTable('Amounts').table.tableRef).to.equal('A1:A4');
    expect(rws.autoFilter).to.equal('A2:A4');
  });
});

const {expect} = require('chai');

const ExcelJS = verquire('exceljs');

describe('Workbook structural row insertion', () => {
  function createWorkbook() {
    const wb = new ExcelJS.Workbook();
    const inputs = wb.addWorksheet('Inputs');
    const summary = wb.addWorksheet('Summary');

    inputs.getCell('A2').value = 100;
    inputs.getCell('A3').value = {formula: 'A2*2', result: 200};
    inputs.getCell('A4').value = {formula: 'SUM(A2:A3)', result: 300};

    summary.getCell('B2').value = {formula: 'Inputs!A3+1', result: 201};
    summary.getCell('B3').value = {formula: 'Inputs!$A$3+1', result: 201};
    summary.getCell('B4').value = {formula: 'SUM(Inputs!A2:A3)', result: 300};
    summary.getCell('B5').value = {
      formula: 'IF(Inputs!A3>0,"Inputs!A3 ok","")',
      result: 'Inputs!A3 ok',
    };

    wb.definedNames.add('Inputs!$A$3', 'ActiveAmount');
    wb.definedNames.add('Inputs!$A$2:$A$3', 'Amounts');

    inputs.dataValidations.add('B3', {
      type: 'list',
      allowBlank: true,
      formulae: ['$D$2:$D$3'],
    });

    inputs.addConditionalFormatting({
      ref: 'A2:A3',
      rules: [{type: 'expression', formulae: ['A2>50'], priority: 1}],
    });

    return {wb, inputs, summary};
  }

  describe('in-memory', () => {
    let state;
    beforeEach(() => {
      state = createWorkbook();
      state.wb.structuralInsertRows('Inputs', 3, 1);
    });

    it('moves cell identities to their new rows', () => {
      const {inputs} = state;
      expect(inputs.getCell('A4').formula).to.equal('A2*2');
      expect(inputs.getCell('A5').formula).to.equal('SUM(A2:A4)');
      expect(inputs.getCell('A3').value).to.equal(null);
    });

    it('updates same-sheet formulas including absolute rows', () => {
      const {summary} = state;
      expect(summary.getCell('B2').formula).to.equal('Inputs!A4+1');
      expect(summary.getCell('B3').formula).to.equal('Inputs!$A$4+1');
      expect(summary.getCell('B4').formula).to.equal('SUM(Inputs!A2:A4)');
    });

    it('leaves address-like text inside string literals alone', () => {
      expect(state.summary.getCell('B5').formula).to.equal(
        'IF(Inputs!A4>0,"Inputs!A3 ok","")'
      );
    });

    it('moves single-cell defined names', () => {
      const model = JSON.stringify(state.wb.definedNames.model);
      expect(model).to.contain('Inputs!$A$4');
      expect(model).not.to.contain('Inputs!$A$3');
    });

    it('expands crossing defined-name ranges instead of splitting them', () => {
      const model = JSON.stringify(state.wb.definedNames.model);
      expect(model).to.contain('Inputs!$A$2:$A$4');
    });

    it('moves data validations with their cells and rewrites their formulae', () => {
      const {dataValidations} = state.inputs;
      expect(dataValidations.model.B3).to.equal(undefined);
      expect(dataValidations.model.B4).to.have.property('type', 'list');
      expect(dataValidations.model.B4.formulae).to.deep.equal(['$D$2:$D$4']);
    });

    it('expands conditional formatting range and keeps rule formula anchored', () => {
      const [cf] = state.inputs.conditionalFormattings;
      expect(cf.ref).to.equal('A2:A4');
      expect(cf.rules[0].formulae).to.deep.equal(['A2>50']);
    });

    it('does not change a sheet that references an unrelated space-named sheet', () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('Inputs');
      const other = wb.addWorksheet('Detail Sheet');
      const summary = wb.addWorksheet('Summary');
      summary.getCell('A1').value = {formula: '\'Detail Sheet\'!A3', result: 0};
      wb.structuralInsertRows('Inputs', 3, 1);
      expect(summary.getCell('A1').formula).to.equal('\'Detail Sheet\'!A3');
      expect(other.name).to.equal('Detail Sheet');
    });

    it('leaves formulas that only reference other sheets untouched', () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('Inputs');
      const a = wb.addWorksheet('A');
      wb.addWorksheet('B');
      a.getCell('A1').value = {formula: 'B!A3+1', result: 0};
      wb.structuralInsertRows('Inputs', 3, 1);
      expect(a.getCell('A1').formula).to.equal('B!A3+1');
    });
  });

  describe('worksheet entry point', () => {
    it('is available on the worksheet and returns new rows', () => {
      const {wb, inputs} = createWorkbook();
      const rows = inputs.structuralInsertRows(3, 2);
      expect(rows).to.have.length(2);
      expect(wb.getWorksheet('Summary').getCell('B2').formula).to.equal('Inputs!A5+1');
    });
  });

  describe('failure handling', () => {
    it('throws on INDIRECT and leaves the workbook untouched', () => {
      const {wb, summary} = createWorkbook();
      summary.getCell('B9').value = {formula: 'INDIRECT("Inputs!A"&3)', result: 1};
      expect(() => wb.structuralInsertRows('Inputs', 3, 1)).to.throw(/INDIRECT/);
      expect(summary.getCell('B2').formula).to.equal('Inputs!A3+1');
      expect(wb.definedNames.model.some(n => n.ranges.includes('Inputs!$A$3'))).to.equal(true);
    });

    it('throws on a 3D reference spanning the edited sheet', () => {
      const {wb} = createWorkbook();
      const summary = wb.getWorksheet('Summary');
      summary.getCell('B10').value = {formula: 'SUM(Inputs:Summary!A3)', result: 0};
      expect(() => wb.structuralInsertRows('Inputs', 3, 1)).to.throw(/3D/);
    });

    it('rejects an invalid start row', () => {
      const {wb} = createWorkbook();
      expect(() => wb.structuralInsertRows('Inputs', 0, 1)).to.throw();
    });
  });

  describe('inheritance options', () => {
    it('does not expand a conditional formatting range when inheritValidations=false', () => {
      const {wb, inputs} = createWorkbook();
      wb.structuralInsertRows('Inputs', 3, 1, {inheritValidations: false});
      expect(inputs.conditionalFormattings[0].ref).to.equal('A2:A3');
    });
  });

  describe('xlsx round trip', () => {
    let loaded;
    beforeEach(async function() {
      this.timeout(5000);
      const state = createWorkbook();
      state.wb.structuralInsertRows('Inputs', 3, 1);

      const wb2 = new ExcelJS.Workbook();
      const buf = await state.wb.xlsx.writeBuffer();
      await wb2.xlsx.load(buf);
      loaded = {
        inputs: wb2.getWorksheet('Inputs'),
        summary: wb2.getWorksheet('Summary'),
        definedNames: wb2.definedNames.model,
      };
    });

    it('preserves moved formulas', () => {
      expect(loaded.inputs.getCell('A4').formula).to.equal('A2*2');
      expect(loaded.inputs.getCell('A5').formula).to.equal('SUM(A2:A4)');
      expect(loaded.summary.getCell('B2').formula).to.equal('Inputs!A4+1');
      expect(loaded.summary.getCell('B3').formula).to.equal('Inputs!$A$4+1');
    });

    it('preserves defined names', () => {
      const model = JSON.stringify(loaded.definedNames);
      expect(model).to.contain('Inputs!$A$4');
      expect(model).to.contain('Inputs!$A$2:$A$4');
    });

    it('preserves data validations', () => {
      const dv = loaded.inputs.dataValidations.model;
      expect(dv.B3).to.equal(undefined);
      expect(dv.B4.formulae).to.deep.equal(['$D$2:$D$4']);
    });

    it('preserves conditional formatting range', () => {
      expect(loaded.inputs.conditionalFormattings[0].ref).to.equal('A2:A4');
    });
  });

  it('does not alter existing spliceRows semantics', () => {
    const {wb, inputs, summary} = createWorkbook();
    inputs.spliceRows(3, 0, []);
    // cross-sheet formulas are untouched by the legacy operation
    expect(summary.getCell('B2').formula).to.equal('Inputs!A3+1');
    expect(wb.definedNames.model.some(n => n.ranges.includes('Inputs!$A$4'))).to.equal(true);
  });
});

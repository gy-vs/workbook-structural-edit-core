const colCache = verquire('utils/col-cache');
const {adjustFormula} = verquire('utils/formula-rewriter');

describe('formula-rewriter', () => {
  const target = 'Inputs';

  function shift(formula, options = {}) {
    return adjustFormula(
      formula,
      options.target || target,
      options.at || 3,
      options.count || 1,
      colCache,
      {ownSheet: options.ownSheet !== false, is3DCovering: options.is3DCovering || (() => false)}
    );
  }

  describe('cross-sheet references', () => {
    it('shifts references to the edited sheet', () => {
      expect(shift('Inputs!A3+1', {ownSheet: false})).to.equal('Inputs!A4+1');
    });

    it('leaves references to other sheets alone', () => {
      expect(shift('Summary!A3+1', {ownSheet: false})).to.equal('Summary!A3+1');
      expect(shift('Summary!A3+1', {ownSheet: true})).to.equal('Summary!A3+1');
    });

    it('handles quoted sheet names with spaces', () => {
      expect(shift('\'My Sheet\'!A3+1', {target: 'My Sheet', ownSheet: false})).to.equal(
        '\'My Sheet\'!A4+1'
      );
      expect(
        shift('\'Other Sheet\'!A3+1', {target: 'My Sheet', ownSheet: false})
      ).to.equal('\'Other Sheet\'!A3+1');
    });

    it('respects absolute row markers', () => {
      expect(shift('Inputs!$A$3', {ownSheet: false})).to.equal('Inputs!$A$4');
      expect(shift('Inputs!A$3', {ownSheet: false})).to.equal('Inputs!A$4');
    });

    it('does not touch rows before the insert position', () => {
      expect(shift('Inputs!A1+Inputs!A2', {ownSheet: false})).to.equal('Inputs!A1+Inputs!A2');
    });
  });

  describe('same-sheet references', () => {
    it('shifts bare references on the edited sheet', () => {
      expect(shift('A3+1')).to.equal('A4+1');
      expect(shift('A$3+1')).to.equal('A$4+1');
    });

    it('shifts both ends of a contiguous range independently', () => {
      expect(shift('SUM(A1:A10)')).to.equal('SUM(A1:A11)');
      expect(shift('SUM(A5:A10)')).to.equal('SUM(A6:A11)');
      expect(shift('SUM(A1:A2)')).to.equal('SUM(A1:A2)');
    });

    it('leaves bare references alone on other sheets', () => {
      expect(shift('A3+1', {ownSheet: false})).to.equal('A3+1');
    });

    it('leaves whole-column references alone', () => {
      expect(shift('SUM(A:A)')).to.equal('SUM(A:A)');
      expect(shift('SUM(Inputs!A:A)', {ownSheet: false})).to.equal('SUM(Inputs!A:A)');
    });

    it('shifts whole-row references', () => {
      // insert through the span: original rows + inserted row
      expect(shift('SUM(1:10)')).to.equal('SUM(1:11)');
      // insert after the span: unchanged
      expect(shift('$1:$2')).to.equal('$1:$2');
      // insert at the first row: span moves with its content
      expect(shift('SUM(3:5)')).to.equal('SUM(4:6)');
    });
  });

  describe('scanning safety', () => {
    it('never treats quoted strings as references', () => {
      expect(shift('"A3"')).to.equal('"A3"');
      expect(shift('IF(A2="A3",A3,0)')).to.equal('IF(A2="A3",A4,0)');
      expect(shift('"Inputs!A3"', {ownSheet: false})).to.equal('"Inputs!A3"');
    });

    it('does not confuse function names and identifiers with references', () => {
      expect(shift('LOG10(A3)')).to.equal('LOG10(A4)');
      expect(shift('100+A3')).to.equal('100+A4');
    });

    it('leaves error and external references alone', () => {
      expect(shift('#REF!')).to.equal('#REF!');
      expect(shift('[1]Inputs!A3', {ownSheet: false})).to.equal('[1]Inputs!A3');
    });
  });

  describe('multiple inserted rows', () => {
    it('shifts by the insert count', () => {
      expect(shift('A3', {at: 2, count: 3})).to.equal('A6');
      expect(shift('SUM(A1:A5)', {at: 3, count: 2})).to.equal('SUM(A1:A7)');
    });
  });

  describe('unsupported references', () => {
    it('throws on R1C1 references into the edited sheet', () => {
      expect(() => shift('R[-1]C')).to.throw();
      expect(() => shift('R1C2')).to.throw();
      expect(() => shift('Inputs!R[1]C2', {ownSheet: false})).to.throw();
    });

    it('does not throw on R1C1 references to other sheets', () => {
      expect(shift('Other!R[-1]C', {ownSheet: false})).to.equal('Other!R[-1]C');
    });

    it('throws on 3D spans that cover the edited sheet', () => {
      expect(() =>
        shift('SUM(Sheet1:Sheet2!A3)', {
          ownSheet: false,
          is3DCovering: () => true,
        })
      ).to.throw();
    });

    it('leaves 3D spans that do not cover the edited sheet alone', () => {
      expect(
        shift('SUM(SheetA:SheetB!A3)', {ownSheet: false, is3DCovering: () => false})
      ).to.equal('SUM(SheetA:SheetB!A3)');
    });
  });
});

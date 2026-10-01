const {
  transformFormula,
  transformRuleFormula,
  transformSqref,
} = verquire('utils/structure-transform');

const EDIT = {sheetName: 'Inputs', start: 3, count: 1};
const SHEETS = ['Inputs', 'Summary', 'My Sheet'];

describe('structure-transform', () => {
  describe('transformFormula - same sheet identities', () => {
    const expectations = [
      ['A2', 'A2'],
      ['A3', 'A4'],
      ['$A$3', '$A$4'],
      ['A3+1', 'A4+1'],
      ['A3+B3+C2', 'A4+B4+C2'],
      // ranges crossing the insert expand; ranges before stay put
      ['SUM(A2:A3)', 'SUM(A2:A4)'],
      ['SUM(A1:A2)', 'SUM(A1:A2)'],
      ['SUM(A5:A6)', 'SUM(A6:A7)'],
      ['SUM($A$2:$A$3)', 'SUM($A$2:$A$4)'],
      // whole rows and whole columns
      ['SUM(1:5)', 'SUM(1:6)'],
      ['SUM(5:6)', 'SUM(6:7)'],
      ['SUM(A:A)', 'SUM(A:A)'],
      // function-name-looking tokens are not references
      ['LOG10(A3)', 'LOG10(A4)'],
      ['MY_NAME*A3', 'MY_NAME*A4'],
      ['10+A3', '10+A4'],
    ];
    expectations.forEach(([input, expected]) => {
      it(input, () => {
        expect(transformFormula(input, EDIT, 'Inputs', SHEETS)).to.equal(expected);
      });
    });
  });

  describe('transformFormula - cross sheet', () => {
    const expectations = [
      ['Inputs!A3+1', 'Inputs!A4+1', 'Summary'],
      ['\'Inputs\'!A3', '\'Inputs\'!A4', 'Summary'],
      ['Inputs!A2:A3', 'Inputs!A2:A4', 'Summary'],
      ['Summary!A3', 'Summary!A3', 'Inputs'],
      ['\'My Sheet\'!A3', '\'My Sheet\'!A3', 'Inputs'],
    ];
    expectations.forEach(([input, expected, current]) => {
      it(`${input} on ${current}`, () => {
        expect(transformFormula(input, EDIT, current, SHEETS)).to.equal(expected);
      });
    });
  });

  describe('transformFormula - string literals', () => {
    [
      ['CONCATENATE("A3",A3)', 'CONCATENATE("A3",A4)'],
      ['"Inputs!A3"', '"Inputs!A3"'],
      ['IF(A3="A3",1,2)', 'IF(A4="A3",1,2)'],
      ['"A""3"A3', '"A""3"A4'],
    ].forEach(([input, expected]) => {
      it(input, () => {
        expect(transformFormula(input, EDIT, 'Inputs', SHEETS)).to.equal(expected);
      });
    });
  });

  describe('transformFormula - unsupported constructs', () => {
    it('rejects 3D references spanning the edited sheet', () => {
      expect(() =>
        transformFormula('SUM(Alpha:Beta!A3)', EDIT, 'Alpha', ['Alpha', 'Inputs', 'Beta'])
      ).to.throw(/3D/);
    });
    it('allows 3D references not spanning the edited sheet', () => {
      expect(
        transformFormula('SUM(Alpha:Beta!A3)', EDIT, 'Alpha', ['Alpha', 'Beta', 'Inputs'])
      ).to.equal('SUM(Alpha:Beta!A3)');
    });
  });

  describe('transformRuleFormula', () => {
    it('relative refs stay put when anchor is above an in-range insertion', () => {
      expect(transformRuleFormula('A2>50', EDIT, 'Inputs', SHEETS, 0)).to.equal('A2>50');
    });
    it('absolute refs follow the moved cell even for an in-range insert', () => {
      expect(transformRuleFormula('A$3>50', EDIT, 'Inputs', SHEETS, 0)).to.equal('A$4>50');
    });
    it('relative refs slide with an anchor above the insertion', () => {
      expect(
        transformRuleFormula('A2>50', {sheetName: 'Inputs', start: 1, count: 1}, 'Inputs', SHEETS, 1)
      ).to.equal('A3>50');
    });
    it('cross-sheet refs are always identities', () => {
      expect(transformRuleFormula('Inputs!A3>1', EDIT, 'Summary', SHEETS, 0)).to.equal(
        'Inputs!A4>1'
      );
    });
  });

  describe('transformSqref', () => {
    [
      ['A2:A3', 'A2:A4'],
      ['B3', 'B4'],
      ['A1 B3 C5:C6', 'A1 B4 C6:C7'],
      ['A1:A2', 'A1:A2'],
    ].forEach(([input, expected]) => {
      it(`${input} -> ${expected}`, () => {
        expect(transformSqref(input, EDIT)).to.equal(expected);
      });
    });
  });
});

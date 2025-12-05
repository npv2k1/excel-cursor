import { Workbook } from 'exceljs';
import { ExcelCursor } from '../src/core/ExcelCursor';

describe('Formula Dependency Issue', () => {
  let workbook: Workbook;
  let cursor: ExcelCursor;

  beforeEach(() => {
    workbook = new Workbook();
    cursor = new ExcelCursor({
      workbook,
      sheetName: 'TestSheet',
    });
  });

  it('should handle formulas that reference other formula cells', () => {
    // Set up data and formulas
    cursor.move('A1').setData(10);         // A1 = 10 (raw value)
    cursor.move('B1').setFormula('=A1*2'); // B1 = A1*2 = 20 (formula)
    cursor.move('C1').setFormula('=B1+5'); // C1 = B1+5 = 25 (formula depending on another formula)

    // These should work correctly
    expect(cursor.getCalculatedValue('B1')).toBe(20);
    expect(cursor.getCalculatedValue('C1')).toBe(25);
  });

  it('should handle SUM with mixed formula and value cells', () => {
    cursor.move('A1').setData(10);         // A1 = 10 (raw value)
    cursor.move('B1').setFormula('=A1*2'); // B1 = 20 (formula)
    cursor.move('C1').setData(5);          // C1 = 5 (raw value)
    
    // This should calculate: 10 + 20 + 5 = 35
    // But currently it might fail because B1 is a formula without cached result
    const result = cursor.calculateFormula('SUM(A1:C1)');
    console.log('SUM result:', result);
    expect(result.error).toBeNull();
    expect(result.result).toBe(35);
  });

  it('should handle range with all formula cells', () => {
    cursor.move('A2').setFormula('=10*3');      // A2 = 30
    cursor.move('B2').setFormula('=15*2');      // B2 = 30
    
    // This should calculate SUM of two formula cells
    const result = cursor.calculateFormula('SUM(A2:B2)');
    console.log('SUM of formulas result:', result);
    expect(result.error).toBeNull();
    expect(result.result).toBe(60);
  });
});
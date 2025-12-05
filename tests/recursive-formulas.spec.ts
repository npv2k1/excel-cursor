import { Workbook } from 'exceljs';
import { ExcelCursor } from '../src/core/ExcelCursor';

describe('Recursive Formula Calculation Test', () => {
  let workbook: Workbook;
  let cursor: ExcelCursor;

  beforeEach(() => {
    workbook = new Workbook();
    cursor = new ExcelCursor({
      workbook,
      sheetName: 'TestSheet',
    });
  });

  it('should handle deep recursive formula chains', () => {
    // Create a deep chain of formulas
    cursor.move('A1').setData(10);              // Base value
    cursor.move('A2').setFormula('=A1*2');      // A2 = 20 (depends on A1)
    cursor.move('A3').setFormula('=A2*2');      // A3 = 40 (depends on A2)
    cursor.move('A4').setFormula('=A3*2');      // A4 = 80 (depends on A3)
    cursor.move('A5').setFormula('=A4*2');      // A5 = 160 (depends on A4)
    cursor.move('A6').setFormula('=SUM(A1:A5)'); // A6 = 10+20+40+80+160 = 310

    // Test that all formulas calculate correctly
    expect(cursor.getCalculatedValue('A2')).toBe(20);
    expect(cursor.getCalculatedValue('A3')).toBe(40);
    expect(cursor.getCalculatedValue('A4')).toBe(80);
    expect(cursor.getCalculatedValue('A5')).toBe(160);
    expect(cursor.getCalculatedValue('A6')).toBe(310);
  });

  it('should handle complex nested formulas in ranges', () => {
    // Set up complex nested formulas
    cursor.move('B1').setFormula('=5*3');       // B1 = 15
    cursor.move('B2').setFormula('=B1+5');      // B2 = 20 (depends on B1)
    cursor.move('B3').setFormula('=B2*2');      // B3 = 40 (depends on B2)
    cursor.move('B4').setFormula('=B3/2');      // B4 = 20 (depends on B3)
    
    // Test SUM of all formula cells
    const result = cursor.calculateFormula('SUM(B1:B4)');
    expect(result.error).toBeNull();
    expect(result.result).toBe(95); // 15 + 20 + 40 + 20 = 95
  });

  it('should handle circular reference detection', () => {
    // This might cause infinite recursion if not handled properly
    cursor.move('C1').setFormula('=C2+1');
    cursor.move('C2').setFormula('=C1+1');
    
    // The engine should handle this gracefully, not crash
    const result1 = cursor.getCalculatedValue('C1');
    const result2 = cursor.getCalculatedValue('C2');
    
    // We expect some kind of error or fallback value, not a crash
    console.log('C1 result:', result1);
    console.log('C2 result:', result2);
  });

  it('should ensure all formulas are fully calculated', () => {
    // Test the scenario where we have multiple levels of dependencies
    cursor.move('D1').setData(100);
    cursor.move('D2').setFormula('=D1*0.1');    // D2 = 10
    cursor.move('D3').setFormula('=D2+D1');     // D3 = 110 (depends on both D1 and D2)
    cursor.move('D4').setFormula('=D3/2');      // D4 = 55 (depends on D3)
    cursor.move('D5').setFormula('=AVERAGE(D1:D4)'); // Average of all
    
    const avgResult = cursor.getCalculatedValue('D5');
    const expected = (100 + 10 + 110 + 55) / 4; // 275/4 = 68.75
    expect(avgResult).toBe(expected);
    
    // Verify all intermediate values are correctly calculated and cached
    expect(cursor.processFormulaCell('D2').hasResult).toBe(true);
    expect(cursor.processFormulaCell('D3').hasResult).toBe(true);
    expect(cursor.processFormulaCell('D4').hasResult).toBe(true);
    expect(cursor.processFormulaCell('D5').hasResult).toBe(true);
  });
});
import { Workbook } from 'exceljs';
import { ExcelCursor } from '../src/index';

// Demonstrate the enhanced recursive formula calculation
async function recursiveFormulaExample() {
  const workbook = new Workbook();
  const cursor = new ExcelCursor({
    workbook,
    sheetName: 'Recursive Formulas',
  });

  console.log('=== Enhanced Recursive Formula Calculation ===');

  // Create a complex dependency chain
  cursor
    .move('A1').setData('Base Values')
    .nextCol().setData(100)      // B1 = 100
    .nextCol().setData(50);      // C1 = 50

  // Create formulas that depend on each other in a chain
  cursor
    .move('A2').setData('Level 1:')
    .nextCol().setFormula('=B1*0.1')    // B2 = B1 * 0.1 = 10
    .nextCol().setFormula('=C1*0.2');   // C2 = C1 * 0.2 = 10

  cursor
    .move('A3').setData('Level 2:')
    .nextCol().setFormula('=B2+C2')     // B3 = B2 + C2 = 20 (depends on B2 and C2)
    .nextCol().setFormula('=B3*2');     // C3 = B3 * 2 = 40 (depends on B3)

  cursor
    .move('A4').setData('Level 3:')
    .nextCol().setFormula('=SUM(B1:C3)') // B4 = SUM of all above
    .nextCol().setFormula('=AVERAGE(B1:C3)'); // C4 = AVERAGE of all above

  // Demonstrate that calculating one formula recursively calculates all dependencies
  console.log('\n1. Before any calculations - checking cache status:');
  console.log('B2 (Level 1) has result:', cursor.processFormulaCell('B2').hasResult);
  console.log('C2 (Level 1) has result:', cursor.processFormulaCell('C2').hasResult);
  console.log('B3 (Level 2) has result:', cursor.processFormulaCell('B3').hasResult);
  console.log('C3 (Level 2) has result:', cursor.processFormulaCell('C3').hasResult);
  console.log('B4 (Level 3) has result:', cursor.processFormulaCell('B4').hasResult);
  console.log('C4 (Level 3) has result:', cursor.processFormulaCell('C4').hasResult);

  console.log('\n2. Calculating B4 (which depends on all other formulas):');
  const b4Result = cursor.getCalculatedValue('B4');
  console.log('B4 calculated value:', b4Result);

  console.log('\n3. After calculating B4 - all dependencies should now be cached:');
  console.log('B2 (Level 1) has result:', cursor.processFormulaCell('B2').hasResult, '- Value:', cursor.processFormulaCell('B2').result);
  console.log('C2 (Level 1) has result:', cursor.processFormulaCell('C2').hasResult, '- Value:', cursor.processFormulaCell('C2').result);
  console.log('B3 (Level 2) has result:', cursor.processFormulaCell('B3').hasResult, '- Value:', cursor.processFormulaCell('B3').result);
  console.log('C3 (Level 2) has result:', cursor.processFormulaCell('C3').hasResult, '- Value:', cursor.processFormulaCell('C3').result);
  console.log('B4 (Level 3) has result:', cursor.processFormulaCell('B4').hasResult, '- Value:', cursor.processFormulaCell('B4').result);

  console.log('\n4. Now calculating C4 should be fast (uses cached dependencies):');
  const c4Result = cursor.getCalculatedValue('C4');
  console.log('C4 calculated value:', c4Result);
  console.log('C4 has result:', cursor.processFormulaCell('C4').hasResult, '- Value:', cursor.processFormulaCell('C4').result);

  // Demonstrate circular reference handling
  console.log('\n5. Testing circular reference handling:');
  cursor.move('D1').setFormula('=D2+1');
  cursor.move('D2').setFormula('=D1+1');
  
  const d1Result = cursor.getCalculatedValue('D1');
  const d2Result = cursor.getCalculatedValue('D2');
  console.log('D1 (circular) result:', d1Result);
  console.log('D2 (circular) result:', d2Result);

  // Add verification data to the spreadsheet
  cursor
    .move('A6').setData('Verification:')
    .nextRow().setData('B2 calculated:').nextCol().setData(cursor.getCalculatedValue('B2'))
    .nextRow().setData('C2 calculated:').nextCol().setData(cursor.getCalculatedValue('C2'))
    .nextRow().setData('B3 calculated:').nextCol().setData(cursor.getCalculatedValue('B3'))
    .nextRow().setData('C3 calculated:').nextCol().setData(cursor.getCalculatedValue('C3'))
    .nextRow().setData('B4 calculated:').nextCol().setData(cursor.getCalculatedValue('B4'))
    .nextRow().setData('C4 calculated:').nextCol().setData(cursor.getCalculatedValue('C4'));

  // Save the file
  await workbook.xlsx.writeFile('./result/recursive-formulas.xlsx');
  console.log('\n✅ Recursive formula example saved to ./result/recursive-formulas.xlsx');

  console.log('\n=== Summary ===');
  console.log('✅ Recursive calculation: When calculating a formula, all its dependencies are automatically calculated and cached');
  console.log('✅ Circular reference detection: Prevents infinite recursion');
  console.log('✅ Performance optimization: Calculated results are cached to avoid recalculation');
  console.log('✅ Deep dependency chains: Handles complex multi-level formula dependencies');
}

recursiveFormulaExample().catch(console.error);
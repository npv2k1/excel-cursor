import { Workbook } from 'exceljs';
import { ExcelCursor } from '../src/index';

// Test the fix for formula dependency calculation
async function testFormulaDependency() {
  const workbook = new Workbook();
  const cursor = new ExcelCursor({
    workbook,
    sheetName: 'Formula Dependency Test',
  });

  console.log('=== Testing Formula Dependency Fix ===');

  // Create a chain of dependent formulas
  cursor
    .move('A1').setData(100)              // Base value
    .move('B1').setFormula('=A1*0.1')     // 10% of A1 = 10
    .move('C1').setFormula('=B1+5')       // B1 + 5 = 15
    .move('D1').setFormula('=C1*2');      // C1 * 2 = 30

  console.log('Chain of formulas:');
  console.log('A1 (base):', cursor.getCellValue('A1'));
  console.log('B1 (=A1*0.1):', cursor.getCalculatedValue('B1'));
  console.log('C1 (=B1+5):', cursor.getCalculatedValue('C1'));  
  console.log('D1 (=C1*2):', cursor.getCalculatedValue('D1'));

  // Test SUM with mixed formula cells
  cursor.move('E1').setFormula('=SUM(A1:D1)'); // Should be 100+10+15+30 = 155
  console.log('E1 (=SUM(A1:D1)):', cursor.getCalculatedValue('E1'));

  // Test complex scenario
  cursor
    .move('A2').setFormula('=10*5')       // 50
    .move('B2').setFormula('=A2/2')       // 25 
    .move('C2').setFormula('=B2+A2')      // 75
    .move('D2').setFormula('=SUM(A2:C2)'); // 150

  console.log('\nComplex formula chain:');
  console.log('A2 (=10*5):', cursor.getCalculatedValue('A2'));
  console.log('B2 (=A2/2):', cursor.getCalculatedValue('B2'));
  console.log('C2 (=B2+A2):', cursor.getCalculatedValue('C2'));
  console.log('D2 (=SUM(A2:C2)):', cursor.getCalculatedValue('D2'));

  // Verify that results are cached
  console.log('\nVerifying cached results:');
  const b1Info = cursor.processFormulaCell('B1');
  const c1Info = cursor.processFormulaCell('C1');
  console.log('B1 has cached result:', b1Info.hasResult, 'Value:', b1Info.result);
  console.log('C1 has cached result:', c1Info.hasResult, 'Value:', c1Info.result);

  // Save file
  await workbook.xlsx.writeFile('./result/formula-dependency-test.xlsx');
  console.log('\n✅ Formula dependency test saved to ./result/formula-dependency-test.xlsx');
}

testFormulaDependency().catch(console.error);
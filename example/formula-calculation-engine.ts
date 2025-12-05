import { Workbook } from 'exceljs';
import { ExcelCursor } from '../src/index';

// Formula calculation engine example
async function formulaCalculationExample() {
  const workbook = new Workbook();
  const cursor = new ExcelCursor({
    workbook,
    sheetName: 'Formula Calculation',
  });

  console.log('Creating formula calculation engine example...');

  // Set up sample data
  cursor
    .move('A1').setData('Sales Data')
    .nextRow().setData(1000)  // A2
    .nextRow().setData(1500)  // A3
    .nextRow().setData(800)   // A4
    .nextRow().setData(1200); // A5

  // Create formulas that will be calculated
  cursor
    .move('C1').setData('Calculated Formulas')
    .nextRow().setFormula('=A2*0.1')              // C2: 10% commission
    .nextRow().setFormula('=SUM(A2:A5)')          // C3: Total sales
    .nextRow().setFormula('=AVERAGE(A2:A5)')      // C4: Average sales
    .nextRow().setFormula('=MAX(A2:A5)')          // C5: Best sale
    .nextRow().setFormula('=MIN(A2:A5)')          // C6: Worst sale
    .nextRow().setFormula('=COUNT(A2:A5)')        // C7: Count of sales
    .nextRow().setFormula('=C3*0.05');            // C8: 5% bonus on total

  // Add labels
  cursor
    .move('B1').setData('Descriptions')
    .nextRow().setData('Commission (10%):')
    .nextRow().setData('Total Sales:')
    .nextRow().setData('Average Sale:')
    .nextRow().setData('Best Sale:')
    .nextRow().setData('Worst Sale:')
    .nextRow().setData('Sale Count:')
    .nextRow().setData('Bonus (5% of total):');

  console.log('\n=== Demonstrating Formula Calculation Engine ===');

  // Test direct formula calculation
  console.log('Direct calculation of "100*2":', cursor.calculateFormula('100*2'));
  console.log('Direct calculation of "SUM(A2:A5)":', cursor.calculateFormula('SUM(A2:A5)'));

  // Calculate values for formula cells
  console.log('\n=== Calculated Values ===');
  console.log('C2 (Commission):', cursor.getCalculatedValue('C2'));
  console.log('C3 (Total Sales):', cursor.getCalculatedValue('C3')); 
  console.log('C4 (Average):', cursor.getCalculatedValue('C4'));
  console.log('C5 (Max):', cursor.getCalculatedValue('C5'));
  console.log('C6 (Min):', cursor.getCalculatedValue('C6'));
  console.log('C7 (Count):', cursor.getCalculatedValue('C7'));
  console.log('C8 (Bonus):', cursor.getCalculatedValue('C8'));

  // Show how cached results work
  console.log('\n=== Showing Cached Results ===');
  const formulaInfo = cursor.processFormulaCell('C3');
  console.log('C3 formula info after calculation:', formulaInfo);

  // Add calculated results to spreadsheet for visualization
  cursor
    .move('D1').setData('Calculated Results')
    .nextRow().setData(cursor.getCalculatedValue('C2'))
    .nextRow().setData(cursor.getCalculatedValue('C3'))
    .nextRow().setData(cursor.getCalculatedValue('C4'))
    .nextRow().setData(cursor.getCalculatedValue('C5'))
    .nextRow().setData(cursor.getCalculatedValue('C6'))
    .nextRow().setData(cursor.getCalculatedValue('C7'))
    .nextRow().setData(cursor.getCalculatedValue('C8'));

  // Test bulk calculation
  console.log('\n=== Testing Bulk Formula Calculation ===');
  
  // Create some more formulas
  cursor
    .move('F1').setData('More Formulas')
    .nextRow().setFormula('=A2+A3')
    .nextRow().setFormula('=A4-A5')
    .nextRow().setFormula('=SUM(A2,A4)');

  console.log('Before bulk calculation:');
  console.log('F2 has result:', cursor.processFormulaCell('F2').hasResult);
  console.log('F3 has result:', cursor.processFormulaCell('F3').hasResult);
  console.log('F4 has result:', cursor.processFormulaCell('F4').hasResult);

  // Calculate all formulas at once
  cursor.calculateAllFormulas();

  console.log('After bulk calculation:');
  console.log('F2 result:', cursor.processFormulaCell('F2').result);
  console.log('F3 result:', cursor.processFormulaCell('F3').result);
  console.log('F4 result:', cursor.processFormulaCell('F4').result);

  // Demonstrate error handling
  console.log('\n=== Error Handling ===');
  cursor.move('G1').setFormula('=UNSUPPORTED_FUNCTION(A1)');
  const errorValue = cursor.getCalculatedValue('G1');
  console.log('Result of unsupported function:', errorValue);

  // Save the file
  await workbook.xlsx.writeFile('./result/formula-calculation-engine.xlsx');
  console.log('\n✅ Formula calculation engine example saved to ./result/formula-calculation-engine.xlsx');

  // Display a summary
  console.log('\n=== Summary ===');
  console.log('The formula calculation engine can:');
  console.log('1. Calculate basic arithmetic: +, -, *, /');
  console.log('2. Handle cell references: A1, B2, etc.');
  console.log('3. Process Excel functions: SUM, AVERAGE, MAX, MIN, COUNT, COUNTA');
  console.log('4. Handle cell ranges: A1:A5');
  console.log('5. Cache calculated results automatically');
  console.log('6. Calculate individual formulas or all at once');
  console.log('7. Handle errors gracefully');
}

// Run the example
formulaCalculationExample().catch(console.error);
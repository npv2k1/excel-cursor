import { Workbook } from 'exceljs';
import { ExcelCursor } from '../src/index';

// Formula processing example showing the new utilities
async function formulaProcessingExample() {
  const workbook = new Workbook();
  const cursor = new ExcelCursor({
    workbook,
    sheetName: 'Formula Processing',
  });

  console.log('Creating formula processing example...');

  // Set up some sample data
  cursor
    .move('A1').setData('Numbers')
    .nextRow().setData(100)
    .nextRow().setData(200)
    .nextRow().setData(300);

  // Add some formulas
  cursor
    .move('C1').setData('Formulas')
    .nextRow().setFormula('=A2*2')       // Simple arithmetic
    .nextRow().setFormula('=SUM(A2:A4)') // SUM function
    .nextRow().setFormula('=AVERAGE(A2:A4)'); // AVERAGE function

  // Add some analysis using the new methods
  cursor
    .move('E1').setData('Analysis')
    .nextRow().setData('Is C2 formula?')
    .nextCol().setData(cursor.isFormulaCell('C2'))
    
    .nextRow().goBackToFirstCollumn().nextCol(4)
    .setData('C2 formula:')
    .nextCol().setData(cursor.getFormula('C2') || 'None')
    
    .nextRow().goBackToFirstCollumn().nextCol(4)
    .setData('C2 raw value:')
    .nextCol().setData(JSON.stringify(cursor.getCellValue('C2')));

  // Demonstrate processing different types of cells
  console.log('\n=== Processing Different Cell Types ===');
  
  // Regular cell
  const regularCellInfo = cursor.processFormulaCell('A2');
  console.log('A2 (number):', regularCellInfo);
  
  // Formula cells
  const formulaCell1 = cursor.processFormulaCell('C2');
  console.log('C2 (formula):', formulaCell1);
  
  const formulaCell2 = cursor.processFormulaCell('C3');
  console.log('C3 (formula):', formulaCell2);

  // Simulate having calculated results (normally done by Excel)
  console.log('\n=== Simulating Calculated Results ===');
  const worksheet = cursor.getWorkbook().getWorksheet('Formula Processing');
  worksheet.getCell('C2').value = { formula: '=A2*2', result: 200 };
  worksheet.getCell('C3').value = { formula: '=SUM(A2:A4)', result: 600 };
  
  // Check the processed results with calculated values
  const withResult1 = cursor.processFormulaCell('C2');
  console.log('C2 (with result):', withResult1);
  
  const withResult2 = cursor.processFormulaCell('C3');
  console.log('C3 (with result):', withResult2);

  // Show difference between raw value and processed value
  console.log('\n=== Value Comparison ===');
  console.log('C2 getCellValue():', cursor.getCellValue('C2'));
  console.log('C2 getFormulaCellValue():', cursor.getFormulaCellValue('C2'));

  // Save the file
  await workbook.xlsx.writeFile('./result/formula-processing.xlsx');
  console.log('\n✅ Formula processing example saved to ./result/formula-processing.xlsx');
}

// Run the example
formulaProcessingExample().catch(console.error);
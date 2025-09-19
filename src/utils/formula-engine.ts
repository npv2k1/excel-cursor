import { Cell, Worksheet } from 'exceljs';
import { parseAddress } from '../helpers/excel.helper';

/**
 * Simple formula evaluation engine for basic Excel formulas
 */
export class FormulaEngine {
  private worksheet: Worksheet;

  constructor(worksheet: Worksheet) {
    this.worksheet = worksheet;
  }

  /**
   * Evaluate a formula and return the calculated result
   * @param formula - The formula string to evaluate (e.g., "=A1+B1", "=SUM(A1:A3)")
   * @returns The calculated result or error
   */
  evaluateFormula(formula: string): { result: any; error: string | null } {
    try {
      // Remove the leading = if present
      const cleanFormula = formula.startsWith('=') ? formula.substring(1) : formula;

      // Handle different formula types
      const result = this.parseAndEvaluate(cleanFormula);
      return { result, error: null };
    } catch (error) {
      return { result: null, error: error instanceof Error ? error.message : 'Unknown error' };
    }
  }

  private parseAndEvaluate(formula: string): any {
    // Handle function calls like SUM(A1:A3), AVERAGE(A1,A2,A3), etc.
    const functionMatch = formula.match(/^([A-Z_]+)\(([^)]+)\)$/);
    if (functionMatch) {
      const [, functionName, args] = functionMatch;
      return this.evaluateFunction(functionName, args);
    }

    // Handle simple arithmetic expressions with cell references
    return this.evaluateExpression(formula);
  }

  private evaluateFunction(functionName: string, args: string): any {
    const values = this.parseArguments(args);

    switch (functionName.toUpperCase()) {
      case 'SUM':
        return values.reduce((sum, val) => sum + (this.isNumeric(val) ? Number(val) : 0), 0);

      case 'AVERAGE':
        const numericValues = values.filter(val => this.isNumeric(val)).map(Number);
        return numericValues.length > 0 
          ? numericValues.reduce((sum, val) => sum + val, 0) / numericValues.length 
          : 0;

      case 'MAX':
        const maxValues = values.filter(val => this.isNumeric(val)).map(Number);
        return maxValues.length > 0 ? Math.max(...maxValues) : 0;

      case 'MIN':
        const minValues = values.filter(val => this.isNumeric(val)).map(Number);
        return minValues.length > 0 ? Math.min(...minValues) : 0;

      case 'COUNT':
        return values.filter(val => this.isNumeric(val)).length;

      case 'COUNTA':
        return values.filter(val => val !== null && val !== undefined && val !== '').length;

      default:
        throw new Error(`Unsupported function: ${functionName}`);
    }
  }

  private parseArguments(args: string): any[] {
    const values: any[] = [];
    const argParts = this.splitArguments(args);

    for (const arg of argParts) {
      const trimmedArg = arg.trim();

      // Check if it's a range (e.g., A1:A3)
      if (trimmedArg.includes(':')) {
        values.push(...this.getRangeValues(trimmedArg));
      }
      // Check if it's a cell reference (e.g., A1)
      else if (this.isCellReference(trimmedArg)) {
        values.push(this.getCellValue(trimmedArg));
      }
      // Otherwise it's a literal value
      else {
        values.push(this.parseLiteral(trimmedArg));
      }
    }

    return values;
  }

  private splitArguments(args: string): string[] {
    const parts: string[] = [];
    let current = '';
    let parentheses = 0;

    for (let i = 0; i < args.length; i++) {
      const char = args[i];
      
      if (char === '(') parentheses++;
      else if (char === ')') parentheses--;
      else if (char === ',' && parentheses === 0) {
        parts.push(current.trim());
        current = '';
        continue;
      }
      
      current += char;
    }
    
    if (current.trim()) {
      parts.push(current.trim());
    }

    return parts;
  }

  private evaluateExpression(expression: string): any {
    // Replace cell references with their values
    let processedExpression = expression;
    const cellReferences = expression.match(/[A-Z]+[0-9]+/g) || [];
    
    for (const cellRef of cellReferences) {
      const cellValue = this.getCellValue(cellRef);
      const numericValue = this.isNumeric(cellValue) ? Number(cellValue) : 0;
      processedExpression = processedExpression.replace(cellRef, numericValue.toString());
    }

    // Evaluate the mathematical expression safely
    return this.safeMathEval(processedExpression);
  }

  private safeMathEval(expression: string): number {
    // Remove any non-math characters for security
    const sanitized = expression.replace(/[^0-9+\-*/.() ]/g, '');
    
    try {
      // Use Function constructor for safe evaluation (better than eval)
      return new Function('return ' + sanitized)();
    } catch (error) {
      throw new Error(`Invalid mathematical expression: ${expression}`);
    }
  }

  private getCellValue(cellAddress: string): any {
    try {
      const { row, col } = parseAddress(cellAddress);
      const cell = this.worksheet.getRow(row).getCell(col);
      
      // If cell has a formula, check if it has a cached result
      if (cell.value && typeof cell.value === 'object' && 'formula' in (cell.value as any)) {
        const cellValue = cell.value as any;
        
        // If already has cached result, return it
        if ('result' in cellValue && cellValue.result !== undefined) {
          return cellValue.result;
        }
        
        // If it's a formula without cached result, calculate it recursively
        const formula = cellValue.formula;
        if (formula) {
          const calculation = this.evaluateFormula(formula);
          if (calculation.error === null) {
            // Cache the result in the cell
            cellValue.result = calculation.result;
            return calculation.result;
          }
        }
        
        // If calculation failed, return 0
        return 0;
      }
      
      return cell.value || 0;
    } catch (error) {
      return 0;
    }
  }

  private getRangeValues(range: string): any[] {
    const [startCell, endCell] = range.split(':');
    const startPos = parseAddress(startCell.trim());
    const endPos = parseAddress(endCell.trim());

    const values: any[] = [];

    for (let row = startPos.row; row <= endPos.row; row++) {
      for (let col = startPos.col; col <= endPos.col; col++) {
        const cell = this.worksheet.getRow(row).getCell(col);
        
        // Get the actual value or result, calculating formulas if needed
        let value = cell.value;
        
        // If cell has a formula, check if it has a cached result
        if (value && typeof value === 'object' && 'formula' in (value as any)) {
          const cellValue = value as any;
          
          // If already has cached result, use it
          if ('result' in cellValue && cellValue.result !== undefined) {
            value = cellValue.result;
          } else {
            // If it's a formula without cached result, calculate it recursively
            const formula = cellValue.formula;
            if (formula) {
              const calculation = this.evaluateFormula(formula);
              if (calculation.error === null) {
                // Cache the result in the cell
                cellValue.result = calculation.result;
                value = calculation.result;
              } else {
                value = 0;
              }
            } else {
              value = 0;
            }
          }
        }
        
        values.push(value || 0);
      }
    }

    return values;
  }

  private isCellReference(str: string): boolean {
    return /^[A-Z]+[0-9]+$/.test(str.trim());
  }

  private isNumeric(value: any): boolean {
    return !isNaN(parseFloat(value)) && isFinite(value);
  }

  private parseLiteral(value: string): any {
    // Try to parse as number
    if (this.isNumeric(value)) {
      return Number(value);
    }
    
    // Try to parse as boolean
    if (value.toLowerCase() === 'true') return true;
    if (value.toLowerCase() === 'false') return false;
    
    // Return as string, removing quotes if present
    return value.replace(/^"(.*)"$/, '$1');
  }
}
import { Cell, Column, Row, Workbook, Worksheet } from 'exceljs';

export interface CellPosition {
  row: number;
  col: number;
}

export interface CellFormat {
  font?: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;
    size?: number;
    name?: string;
  };
  fill?: {
    type: string;
    pattern?: string;
    fgColor?: string;
    bgColor?: string;
  };
  border?: {
    top?: { style: string; color: string };
    left?: { style: string; color: string };
    bottom?: { style: string; color: string };
    right?: { style: string; color: string };
  };
  alignment?: {
    vertical?: 'top' | 'middle' | 'bottom';
    horizontal?: 'left' | 'center' | 'right';
    wrapText?: boolean;
  };
  numFmt?: string;
}

export type { Cell, Column, Row, Workbook, Worksheet };

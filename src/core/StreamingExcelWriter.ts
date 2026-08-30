import { stream } from 'exceljs';

import { EXCEL_MAX_ROWS } from '../helpers/excel.helper';

export interface StreamingProgress {
  rowsWritten: number;
  maxRows: number;
}

export interface StreamingExcelWriterOptions {
  filename: string;
  sheetName?: string;
  maxRows?: number;
  signal?: AbortSignal;
  onProgress?: (progress: StreamingProgress) => void;
  useSharedStrings?: boolean;
  useStyles?: boolean;
}

type RowValues = ReadonlyArray<unknown>;
type WriterState = 'open' | 'committing' | 'committed' | 'failed';

/**
 * Append-only XLSX writer for data sets that should not be held in memory.
 * Rows are committed immediately and cannot be read or modified afterwards.
 */
export class StreamingExcelWriter {
  private readonly workbook: stream.xlsx.WorkbookWriter;
  private readonly worksheet: ReturnType<stream.xlsx.WorkbookWriter['addWorksheet']>;
  private readonly maxRows: number;
  private readonly signal?: AbortSignal;
  private readonly onProgress?: (progress: StreamingProgress) => void;
  private rowsWritten = 0;
  private state: WriterState = 'open';
  private commitPromise?: Promise<void>;

  constructor(options: StreamingExcelWriterOptions) {
    if (!options.filename) throw new Error('Streaming writer filename is required');
    const maxRows = options.maxRows ?? EXCEL_MAX_ROWS;
    if (!Number.isInteger(maxRows) || maxRows < 1 || maxRows > EXCEL_MAX_ROWS) {
      throw new Error(`Invalid maxRows: ${maxRows}`);
    }

    this.maxRows = maxRows;
    this.signal = options.signal;
    this.onProgress = options.onProgress;
    this.throwIfAborted();
    this.workbook = new stream.xlsx.WorkbookWriter({
      filename: options.filename,
      useSharedStrings: options.useSharedStrings ?? true,
      useStyles: options.useStyles ?? true,
    });
    this.worksheet = this.workbook.addWorksheet(options.sheetName || 'Sheet1');
  }

  addRow(values: RowValues): this {
    this.assertOpen();
    this.throwIfAborted();
    if (this.rowsWritten >= this.maxRows) {
      throw new Error(`Streaming row limit exceeded: ${this.maxRows}`);
    }

    this.worksheet.addRow(Array.from(values)).commit();
    this.rowsWritten += 1;
    this.onProgress?.({ rowsWritten: this.rowsWritten, maxRows: this.maxRows });
    return this;
  }

  async addRows(rows: Iterable<RowValues> | AsyncIterable<RowValues>): Promise<number> {
    for await (const row of rows) {
      this.addRow(row);
    }
    return this.rowsWritten;
  }

  getRowsWritten(): number {
    return this.rowsWritten;
  }

  commit(): Promise<void> {
    if (this.commitPromise !== undefined) return this.commitPromise;
    this.assertOpen();
    this.state = 'committing';
    this.commitPromise = this.workbook.commit().then(
      () => {
        this.state = 'committed';
      },
      (error) => {
        this.state = 'failed';
        throw error;
      }
    );
    return this.commitPromise;
  }

  private assertOpen(): void {
    if (this.state !== 'open') {
      throw new Error(`Streaming writer is ${this.state}; no more rows can be added`);
    }
  }

  private throwIfAborted(): void {
    if (this.signal?.aborted) {
      const error = new Error('Streaming write aborted');
      error.name = 'AbortError';
      throw error;
    }
  }
}

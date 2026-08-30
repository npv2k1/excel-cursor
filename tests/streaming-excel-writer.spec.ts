import { Workbook } from 'exceljs';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

import { ExcelCursor } from '../src/core/ExcelCursor';
import { StreamingExcelWriter } from '../src/core/StreamingExcelWriter';

describe('StreamingExcelWriter', () => {
  let directory: string;

  beforeEach(() => {
    directory = fs.mkdtempSync(path.join(os.tmpdir(), 'excel-cursor-stream-test-'));
  });

  afterEach(() => {
    fs.rmSync(directory, { recursive: true, force: true });
  });

  it('writes and commits rows that can be reopened', async () => {
    const filename = path.join(directory, 'rows.xlsx');
    const progress: number[] = [];
    const writer = new StreamingExcelWriter({
      filename,
      sheetName: 'Data',
      onProgress: ({ rowsWritten }) => progress.push(rowsWritten),
    });

    writer.addRow(['id', 'name']).addRow([1, 'Ada']);
    await writer.commit();
    await writer.commit();

    const workbook = await new Workbook().xlsx.readFile(filename);
    expect(workbook.getWorksheet('Data')?.getRow(2).values).toEqual([, 1, 'Ada']);
    expect(progress).toEqual([1, 2]);
    expect(writer.getRowsWritten()).toBe(2);
    expect(() => writer.addRow([2, 'Grace'])).toThrow('committed');
  });

  it('consumes AsyncIterable rows with bounded memory semantics', async () => {
    const filename = path.join(directory, 'async.xlsx');
    async function* rows() {
      yield [1];
      await Promise.resolve();
      yield [2];
    }

    const writer = new StreamingExcelWriter({ filename });
    await expect(writer.addRows(rows())).resolves.toBe(2);
    await writer.commit();
    expect(fs.statSync(filename).size).toBeGreaterThan(0);
  });

  it('enforces maxRows before appending another row', async () => {
    const writer = new StreamingExcelWriter({
      filename: path.join(directory, 'limited.xlsx'),
      maxRows: 1,
    });
    writer.addRow(['only']);
    expect(() => writer.addRow(['excess'])).toThrow('row limit exceeded');
    await writer.commit();
  });

  it('honors AbortSignal between rows', async () => {
    const controller = new AbortController();
    const writer = new StreamingExcelWriter({
      filename: path.join(directory, 'aborted.xlsx'),
      signal: controller.signal,
    });
    writer.addRow([1]);
    controller.abort();
    expect(() => writer.addRow([2])).toThrow(expect.objectContaining({ name: 'AbortError' }));
    // Finalizing remains possible after cancellation so file handles are closed
    // and the successfully committed prefix is a valid workbook.
    await writer.commit();
  });

  it('gives ExcelCursor streaming callers an explicit save lifecycle error', async () => {
    const filename = path.join(directory, 'legacy.xlsx');
    const cursor = new ExcelCursor({ isStream: true, filename });
    cursor.addRow(['legacy']);
    await expect(cursor.saveWorkbook(filename)).rejects.toThrow('call commit()');
    await cursor.commit();
    expect(fs.statSync(filename).size).toBeGreaterThan(0);
  });
});

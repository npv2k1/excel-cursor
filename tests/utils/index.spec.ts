import { stream, Workbook } from 'exceljs';
import * as fs from 'fs';

import { createStreamWorkbook, createWorkbook } from '../../src/utils/index';

describe('utils/index', () => {
  describe('createWorkbook', () => {
    it('should create a new Workbook instance', () => {
      const workbook = createWorkbook();
      expect(workbook).toBeInstanceOf(Workbook);
    });

    it('should create a fresh workbook each time', () => {
      const workbook1 = createWorkbook();
      const workbook2 = createWorkbook();
      expect(workbook1).not.toBe(workbook2);
    });

    it('should create workbook with no worksheets by default', () => {
      const workbook = createWorkbook();
      expect(workbook.worksheets).toHaveLength(0);
    });
  });

  describe('createStreamWorkbook', () => {
    it('should create a WorkbookWriter instance', () => {
      const options = {
        filename: '/tmp/test.xlsx',
        stream: fs.createWriteStream('/tmp/test.xlsx'),
        useStyles: true,
        useSharedStrings: true,
      };
      const workbook = createStreamWorkbook(options);
      expect(workbook).toBeInstanceOf(stream.xlsx.WorkbookWriter);
    });

    it('should create workbook with provided options', () => {
      const options = {
        filename: '/tmp/test.xlsx',
        stream: fs.createWriteStream('/tmp/test.xlsx'),
        useStyles: true,
        useSharedStrings: true,
      };
      const workbook = createStreamWorkbook(options);
      expect(workbook).toBeInstanceOf(stream.xlsx.WorkbookWriter);
    });

    it('should create a fresh workbook each time', () => {
      const options = {
        filename: '/tmp/test1.xlsx',
        stream: fs.createWriteStream('/tmp/test1.xlsx'),
        useStyles: true,
        useSharedStrings: true,
      };
      const workbook1 = createStreamWorkbook(options);
      const workbook2 = createStreamWorkbook({
        filename: '/tmp/test2.xlsx',
        stream: fs.createWriteStream('/tmp/test2.xlsx'),
        useStyles: true,
        useSharedStrings: true,
      });
      expect(workbook1).not.toBe(workbook2);
    });
  });
});

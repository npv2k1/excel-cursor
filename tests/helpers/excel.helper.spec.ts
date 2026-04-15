import {
  colLetterToNumber,
  colNumberToLetter,
  parseAddress,
  positionToAddress,
} from '../../src/helpers/excel.helper';

describe('excel.helper', () => {
  describe('colLetterToNumber', () => {
    it('should convert A to 1', () => {
      expect(colLetterToNumber('A')).toBe(1);
    });

    it('should convert Z to 26', () => {
      expect(colLetterToNumber('Z')).toBe(26);
    });

    it('should convert AA to 27', () => {
      expect(colLetterToNumber('AA')).toBe(27);
    });

    it('should convert AZ to 52', () => {
      expect(colLetterToNumber('AZ')).toBe(52);
    });

    it('should convert BA to 53', () => {
      expect(colLetterToNumber('BA')).toBe(53);
    });

    it('should convert ZZ to 702', () => {
      expect(colLetterToNumber('ZZ')).toBe(702);
    });

    it('should convert AAA to 703', () => {
      expect(colLetterToNumber('AAA')).toBe(703);
    });

    it('should convert B to 2', () => {
      expect(colLetterToNumber('B')).toBe(2);
    });

    it('should convert C to 3', () => {
      expect(colLetterToNumber('C')).toBe(3);
    });

    it('should convert M to 13', () => {
      expect(colLetterToNumber('M')).toBe(13);
    });
  });

  describe('colNumberToLetter', () => {
    it('should convert 1 to A', () => {
      expect(colNumberToLetter(1)).toBe('A');
    });

    it('should convert 26 to Z', () => {
      expect(colNumberToLetter(26)).toBe('Z');
    });

    it('should convert 27 to AA', () => {
      expect(colNumberToLetter(27)).toBe('AA');
    });

    it('should convert 52 to AZ', () => {
      expect(colNumberToLetter(52)).toBe('AZ');
    });

    it('should convert 53 to BA', () => {
      expect(colNumberToLetter(53)).toBe('BA');
    });

    it('should convert 702 to ZZ', () => {
      expect(colNumberToLetter(702)).toBe('ZZ');
    });

    it('should convert 703 to AAA', () => {
      expect(colNumberToLetter(703)).toBe('AAA');
    });

    it('should convert 2 to B', () => {
      expect(colNumberToLetter(2)).toBe('B');
    });

    it('should convert 3 to C', () => {
      expect(colNumberToLetter(3)).toBe('C');
    });

    it('should convert 13 to M', () => {
      expect(colNumberToLetter(13)).toBe('M');
    });
  });

  describe('parseAddress', () => {
    it('should parse A1', () => {
      expect(parseAddress('A1')).toEqual({ row: 1, col: 1 });
    });

    it('should parse B2', () => {
      expect(parseAddress('B2')).toEqual({ row: 2, col: 2 });
    });

    it('should parse Z26', () => {
      expect(parseAddress('Z26')).toEqual({ row: 26, col: 26 });
    });

    it('should parse AA100', () => {
      expect(parseAddress('AA100')).toEqual({ row: 100, col: 27 });
    });

    it('should parse C5', () => {
      expect(parseAddress('C5')).toEqual({ row: 5, col: 3 });
    });

    it('should throw error for invalid address without letters', () => {
      expect(() => parseAddress('123')).toThrow('Invalid cell address: 123');
    });

    it('should throw error for invalid address without numbers', () => {
      expect(() => parseAddress('ABC')).toThrow('Invalid cell address: ABC');
    });

    it('should throw error for empty string', () => {
      expect(() => parseAddress('')).toThrow('Invalid cell address: ');
    });

    it('should throw error for address with special characters', () => {
      expect(() => parseAddress('A@1')).toThrow('Invalid cell address: A@1');
    });
  });

  describe('positionToAddress', () => {
    it('should convert (1, 1) to A1', () => {
      expect(positionToAddress(1, 1)).toBe('A1');
    });

    it('should convert (2, 2) to B2', () => {
      expect(positionToAddress(2, 2)).toBe('B2');
    });

    it('should convert (26, 26) to Z26', () => {
      expect(positionToAddress(26, 26)).toBe('Z26');
    });

    it('should convert (100, 27) to AA100', () => {
      expect(positionToAddress(100, 27)).toBe('AA100');
    });

    it('should convert (5, 3) to C5', () => {
      expect(positionToAddress(5, 3)).toBe('C5');
    });

    it('should convert (1, 52) to AZ1', () => {
      expect(positionToAddress(1, 52)).toBe('AZ1');
    });

    it('should convert (1, 703) to AAA1', () => {
      expect(positionToAddress(1, 703)).toBe('AAA1');
    });
  });

  describe('Round trip conversion', () => {
    it('should correctly round trip for single letters', () => {
      for (let i = 1; i <= 26; i++) {
        const letter = colNumberToLetter(i);
        const number = colLetterToNumber(letter);
        expect(number).toBe(i);
      }
    });

    it('should correctly round trip for double letters', () => {
      for (let i = 27; i <= 702; i += 25) {
        const letter = colNumberToLetter(i);
        const number = colLetterToNumber(letter);
        expect(number).toBe(i);
      }
    });

    it('should correctly round trip for address parsing', () => {
      const addresses = ['A1', 'B5', 'Z26', 'AA100', 'ZZ999', 'AAA1'];
      addresses.forEach((address) => {
        const pos = parseAddress(address);
        const result = positionToAddress(pos.row, pos.col);
        expect(result).toBe(address);
      });
    });
  });
});

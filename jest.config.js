const { pathsToModuleNameMapper } = require('ts-jest');
const { compilerOptions } = {
  compilerOptions: {
    baseUrl: './src',
    paths: {
      '@/*': ['src/*'],
    },
  },
};

/** @type {import('jest').Config} */
module.exports = {
  moduleFileExtensions: ['js', 'json', 'ts'],
  // collectCoverage: true,
  coverageProvider: 'v8',
  transform: {
    '^.+\\.(js|ts)$': 'ts-jest',
  },
  collectCoverageFrom: [
    '**/*.{js,jsx,ts,tsx}',
    '!**/*.stories.{js,jsx,ts,tsx}',
    '!**/*.d.ts',
    '!**/node_modules/**',
    '!<rootDir>/out/**',
    '!<rootDir>/.next/**',
    '!<rootDir>/*.config.js',
    '!<rootDir>/coverage/**',
  ],
  testEnvironment: 'node',
  rootDir: './',
  roots: ['<rootDir>/src/', '<rootDir>/tests/'],
  testRegex: '(/tests/.*|(\\.|/)(test|spec))\\.(jsx?|tsx?)$',
  coverageDirectory: './coverage',
  moduleNameMapper: pathsToModuleNameMapper(compilerOptions.paths),
  moduleDirectories: ['node_modules', '<rootDir>'],
};

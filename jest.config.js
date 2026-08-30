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
  watchman: false,
  moduleFileExtensions: ['js', 'json', 'ts'],
  // collectCoverage: true,
  coverageProvider: 'v8',
  transform: {
    '^.+\\.(js|ts)$': 'ts-jest',
  },
  collectCoverageFrom: [
    '<rootDir>/src/core/**/*.ts',
    '<rootDir>/src/helpers/**/*.ts',
    '<rootDir>/src/utils/index.ts',
  ],
  testEnvironment: 'node',
  rootDir: './',
  roots: ['<rootDir>/src/', '<rootDir>/tests/'],
  testRegex: '(/tests/.*|(\\.|/)(test|spec))\\.(jsx?|tsx?)$',
  coverageDirectory: './coverage',
  coverageThreshold: {
    global: {
      branches: 90,
      functions: 95,
      lines: 95,
      statements: 95,
    },
  },
  moduleNameMapper: pathsToModuleNameMapper(compilerOptions.paths),
  moduleDirectories: ['node_modules', '<rootDir>'],
};

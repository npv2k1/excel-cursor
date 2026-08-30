const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { spawnSync } = require('node:child_process');

function run(command, args, options = {}) {
  const result = spawnSync(command, args, { encoding: 'utf8', ...options });
  if (result.status !== 0) {
    throw new Error(`${command} ${args.join(' ')} failed:\n${result.stderr || result.stdout}`);
  }
  return result.stdout;
}

async function main() {
  const projectRoot = path.resolve(__dirname, '..');
  const temporaryDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'excel-cursor-package-smoke-'));
  try {
    const packResult = JSON.parse(
      run(
        'npm',
        ['pack', projectRoot, '--json', '--ignore-scripts', '--pack-destination', temporaryDirectory],
        {
          cwd: projectRoot,
          env: { ...process.env, npm_config_cache: path.join(temporaryDirectory, 'npm-cache') },
        },
      ),
    );
    const tarball = path.join(temporaryDirectory, packResult[0].filename);
    run('tar', ['-xzf', tarball], { cwd: temporaryDirectory });
    const packageRoot = path.join(temporaryDirectory, 'package');
    fs.symlinkSync(path.join(projectRoot, 'node_modules'), path.join(packageRoot, 'node_modules'));
    const manifest = JSON.parse(fs.readFileSync(path.join(packageRoot, 'package.json'), 'utf8'));

    assert.equal(manifest.exports['.'].require, './dist/cjs/index.js');
    assert.equal(manifest.exports['.'].import, './dist/esm/index.js');
    assert.ok(fs.existsSync(path.join(packageRoot, 'dist/esm/index.d.ts')));

    const commonJs = require(packageRoot);
    assert.equal(typeof commonJs.ExcelCursor, 'function');

    const esm = await import(pathToFileURL(path.join(packageRoot, 'dist/esm/index.js')).href);
    assert.equal(typeof esm.ExcelCursor, 'function');
    process.stdout.write('Packed package CJS, ESM, and types smoke tests passed.\n');
  } finally {
    fs.rmSync(temporaryDirectory, { recursive: true, force: true });
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

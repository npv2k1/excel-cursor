const fs = require('node:fs');
const path = require('node:path');

const esmDirectory = path.resolve(__dirname, '..', 'dist', 'esm');

function visit(directory) {
  for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
    const filename = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      visit(filename);
      continue;
    }
    if (!entry.name.endsWith('.js')) continue;

    const source = fs.readFileSync(filename, 'utf8');
    let rewritten = source;
    for (const [packageName, binding] of [
      ['exceljs', 'exceljs'],
      ['lodash', 'lodash'],
    ]) {
      rewritten = rewritten.replace(
        new RegExp(`import \\{([^}]+)\\} from ['\"]${packageName}['\"];?`, 'g'),
        (_match, names) =>
          `import ${binding} from '${packageName}';\nconst { ${names.trim()} } = ${binding};`,
      );
    }
    rewritten = rewritten.replace(
      /(\b(?:from\s+|import\s*\(|export\s+[^'\"]*?from\s+)[\"'])(\.{1,2}\/[^'\"]+)([\"'])/g,
      (match, prefix, specifier, suffix) => {
        if (/\.(?:js|mjs|cjs|json|node)$/.test(specifier)) return match;
        const absoluteTarget = path.resolve(path.dirname(filename), specifier);
        const target = fs.existsSync(`${absoluteTarget}.js`)
          ? `${specifier}.js`
          : `${specifier}/index.js`;
        return `${prefix}${target}${suffix}`;
      },
    );
    fs.writeFileSync(filename, rewritten);
  }
}

visit(esmDirectory);
fs.writeFileSync(
  path.join(esmDirectory, 'package.json'),
  `${JSON.stringify({ type: 'module' }, null, 2)}\n`,
);

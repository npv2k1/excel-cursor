# Contributing

## Development setup

Use Node.js 22 or 24 and the pnpm version declared in `package.json`.

```bash
corepack enable
pnpm install --frozen-lockfile
```

## Before opening a pull request

```bash
pnpm run format:check
pnpm run lint
pnpm run test:ci
pnpm run build
pnpm run package:smoke
```

Add regression tests for behavior changes. Workbook features should normally include a write/reopen assertion, while packaging changes must test the packed tarball rather than importing `dist` directly.

Keep changes focused, preserve backward compatibility unless a major release has been agreed, and document user-visible changes in `CHANGELOG.md`. Do not commit generated `dist`, coverage, result workbooks, credentials, or customer data.

## Security and trust boundaries

Do not add examples that pass user input to formula APIs or expose arbitrary output paths. Security reports belong in GitHub's private vulnerability-reporting flow described in [SECURITY.md](SECURITY.md), not a public issue.

By contributing, you agree that your contribution is licensed under the repository's MIT license.

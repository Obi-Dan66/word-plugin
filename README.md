# word-plugin

Small Word add-in using Office.js (scaffolded from the Office task pane template).

## Setup

Use a compatible Node version (≥ 22.13 or v26):

```bash
nvm use 26.1.0
yarn install
```

## Scripts

| Command | Description |
|---------|-------------|
| `yarn dev-server` | Webpack dev server (HTTPS on port 3000) |
| `yarn start` | Sideload add-in in Word for debugging |
| `yarn build` | Production build |
| `yarn validate` | Validate `manifest.xml` |

Scaffolding tools (`yo`, `generator-office`) are dev dependencies: `yarn yo office`.

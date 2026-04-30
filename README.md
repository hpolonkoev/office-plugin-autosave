# AutoSave for Office

[![Deploy to GitHub Pages](https://github.com/hpolonkoev/office-plugin-autosave/actions/workflows/deploy.yml/badge.svg)](https://github.com/hpolonkoev/office-plugin-autosave/actions/workflows/deploy.yml)
![Version](https://img.shields.io/github/package-json/v/hpolonkoev/office-plugin-autosave)
![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![Office.js](https://img.shields.io/badge/Office.js-Word%20%7C%20Excel%20%7C%20PowerPoint-0078d4)

> A Microsoft Office Web Add-in that silently saves your documents on a configurable schedule — no clicks, no interruptions, no lost work.

**[View product page →](https://hpolonkoev.github.io/office-plugin-autosave/)**

---

## Features

- **Works across Word, Excel & PowerPoint** — one manifest, three apps
- **Runs in the background** — uses Office Shared Runtime so the timer keeps going even when the task pane is closed
- **IT-configurable** — set the default interval, min/max bounds, and whether users can override, via a single JSON file on the server
- **Multi-language** — ships with English, French, and Dutch; drop a new `.json` file on the server to add more, no rebuild needed
- **Protected View aware** — detects read-only documents, disables the timer silently
- **Zero infrastructure** — fully client-side, no backend, no database

---

## For IT Admins

Download [`manifest.xml`](https://hpolonkoev.github.io/office-plugin-autosave/manifest.xml) and deploy via **Microsoft 365 Admin Center → Settings → Integrated apps**.

Configure intervals and user permissions by editing `config/autosave-config.json` on the server — no rebuild required.

→ [Full IT configuration guide](docs/IT_GUIDE.md)

---

## Local Development

**Prerequisites:** Node.js 18+, Microsoft 365 subscription

```bash
npm install
npm start
```

Opens the webpack dev server at `https://localhost:3000`. Load `manifest.xml` from a trusted catalog in Word, Excel, or PowerPoint to test.

### Build for production

```bash
npm run build
```

Output lands in `dist/`. The GitHub Actions workflow builds and deploys to GitHub Pages automatically on every push to `main`.

---

## Tech Stack

- [Office.js](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/javascript-api-for-office) — Office Web Add-in API
- TypeScript + Webpack
- Plain HTML/CSS (no UI framework)
- Custom lightweight i18n module

---

## License

MIT
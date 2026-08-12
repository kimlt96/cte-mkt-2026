# Cao Thắng Marketing Report

Internal weekly marketing report dashboard (Digital / Task List / Planning) for Cao Thắng Hospital, built with React + Vite.

## Features

- **Digital tab** — Budget (Total/Detail), Digital Channels (Summary/Branding Ads), and a LASIK funnel section with automatic conversion-rate calculation.
- **Task List tab** — dynamic task table driven entirely by imported Excel columns, with search/filter and an image report section.
- **Planning tab** — image-based planning cards with Format/BU filters.
- **Import Data** — upload `.xlsx`/`.xls`/`.csv`, auto-detect columns/rows (no hard-coded table shape), with column mapping for the LASIK KPI/Actual funnel fields.
- **Images** — upload, replace, delete, and drag-to-reorder, always rendered at their original aspect ratio (no cropping).
- **Report history** — save/load weekly snapshots (e.g. `W32`, `W33`), with a duplicate-report guard (Update existing / Create new version / Cancel).
- **Download PDF** — uses the browser's native print-to-PDF (`window.print()`) with dedicated print styles.
- **Persistent storage** — all report data and images are saved locally in the browser via IndexedDB (see `src/lib/storage.js`), so data survives refreshes/restarts on the same device/browser.

## Getting started

```bash
npm install
npm run dev
```

The app will be available at `http://localhost:5173`.

## Build for production

```bash
npm run build
npm run preview   # optional: preview the production build locally
```

The build output is written to `dist/`.

## Deploying to Vercel

This project is ready to deploy as-is:

1. Push this folder to a Git repository (GitHub/GitLab/Bitbucket).
2. Import the repo in [Vercel](https://vercel.com/new).
3. Vercel auto-detects the Vite framework (build command `npm run build`, output directory `dist`) — no extra configuration needed. A `vercel.json` is included for clarity and to support client-side routing.
4. Deploy.

Alternatively, using the Vercel CLI:

```bash
npm i -g vercel
vercel
```

## Project structure

```
├── index.html
├── package.json
├── vite.config.js
├── vercel.json
├── public/
│   └── favicon.svg
└── src/
    ├── main.jsx        # React entry point
    ├── App.jsx          # Entire dashboard (all tabs, modals, table/image components)
    ├── styles.css        # All app styling (design tokens, layout, print styles)
    └── lib/
        └── storage.js    # IndexedDB-backed persistence layer
```

## Notes on data storage

The app currently persists everything (report data + images) in the browser's IndexedDB via `src/lib/storage.js`, which exposes a small `get/set/delete/list` contract. This mirrors the original prototype's storage design so it can later be swapped for a real backend (e.g. Supabase/PostgreSQL) by re-implementing those four functions — no changes needed in `App.jsx`.

Because storage is per-browser/per-device, report history will not sync across different computers or browsers unless a backend is added.

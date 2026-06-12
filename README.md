# quran-similarities · المتشابهات اللفظية في القرآن

A simple, **fully static** React app (Vite) for browsing the verbal similarities
(المتشابهات اللفظية) of each surah and downloading a formatted Word (`.docx`) file.
No server or backend — everything runs in the browser.

## How it works

- The 114 surah datasets live in [public/data/](public/data/) and are fetched at
  runtime as static files (`/data/{num}.json`).
- The Word document is built **client-side** with the `docx` library
  (`Packer.toBlob`) — see [src/buildDocx.js](src/buildDocx.js). The old Express
  `/word/:num` route is gone.
- Clicking a surah renders an on-screen preview and downloads the matching `.docx`.

## Develop

```bash
npm install
npm run dev      # local dev server (Vite)
```

## Build / deploy

```bash
npm run build    # outputs a static site to dist/
npm run preview  # preview the production build
```

`dist/` is a plain static bundle — host it on any static host (GitHub Pages,
Netlify, Vercel, S3, etc.) with no backend.

## Updating the data

[fetch_similarities.js](fetch_similarities.js) re-downloads the source JSON into
`public/data/` (run with Node). Only needed when refreshing the dataset.

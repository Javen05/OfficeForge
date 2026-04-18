# Office Forge

Office Forge is a responsive Next.js workspace for editing DOCX, PDF, PPT, and XLSX content in one place.

## What it does

- Imports Office and PDF files into a document library
- Provides type-aware editing surfaces for documents, presentations, spreadsheets, and PDF review notes
- Autosaves workspace state in the browser
- Optimizes the layout for both mobile and desktop screens

## Run locally

```bash
npm install
npm run dev
```

## Build

```bash
npm run build
npm run start
```

## Notes

- DOCX files are converted into editable text on import.
- XLSX files load into an editable grid.
- PPT files open into a slide editor.
- PDF files open into a review surface with notes and preview support.

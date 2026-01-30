# Draftworx Context — Project Context

## Overview

Excel add-in that extracts selected cell data (references, values, formulas) as structured JSON for pasting into AI chat windows. Solves the "typing context is painful" problem.

## Tech Stack

- **Language:** TypeScript
- **API:** Office.js (Excel JavaScript API)
- **Build:** Webpack 5
- **Hosting:** Local dev server (https://localhost:3000)

## File Structure

```
draftworx-context/
├── manifest.xml           # Office add-in manifest
├── package.json           # Dependencies
├── tsconfig.json          # TypeScript config
├── webpack.config.js      # Build config
├── src/
│   └── taskpane/
│       ├── taskpane.html  # UI
│       └── taskpane.ts    # Core logic
└── assets/                # Icons (TODO)
```

## Key APIs

```typescript
// Get selected range
const range = context.workbook.getSelectedRange();
range.load(['address', 'values', 'formulas']);

// Listen for selection changes
sheet.onSelectionChanged.add(handler);
```

## Commands

```bash
npm install      # Install deps
npm run dev      # Start dev server
npm run build    # Production build
npm run sideload # Load into Excel
```

## Roadmap

- [ ] Basic selection → JSON extraction
- [ ] Live update on selection change
- [ ] Copy to clipboard
- [ ] Settings (include formulas, etc.)
- [ ] AI chat integration (Claude API)
- [ ] Production deployment
- [ ] Named range support
- [ ] Table detection
- [ ] Cross-sheet references

## References

- [Office.js Excel API](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/)
- [Manifest XML Schema](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests)

# Draftworx Context

Excel add-in for AI-powered assistance. Extracts selected cell data as structured JSON — perfect for pasting into AI chat windows.

## Features

- 🔴 **Live Updates** — Automatically captures selection changes
- 📋 **One-Click Copy** — JSON to clipboard instantly
- 📊 **Rich Data** — Cell references, values, and formulas
- ⚡ **Lightweight** — Fast, minimal UI that stays out of your way

## JSON Output Format

```json
{
  "selection": "A1:C3",
  "sheet": "Sheet1",
  "timestamp": "2026-01-30T05:00:00.000Z",
  "cells": [
    { "ref": "A1", "value": "Product", "formula": null },
    { "ref": "B1", "value": 100, "formula": null },
    { "ref": "C1", "value": 150, "formula": "=B1*1.5" }
  ]
}
```

## Development

### Prerequisites

- Node.js 18+ (LTS recommended)
- Excel (Windows, Mac, or Web)

### Setup

```bash
# Install dependencies
npm install

# Start dev server (https://localhost:3000)
npm run dev

# Sideload into Excel
npm run sideload
```

### Build

```bash
npm run build
```

## Installation (Production)

1. Get the manifest URL from your deployment
2. In Excel: Insert → Add-ins → Upload My Add-in
3. Select the manifest.xml file

## Tech Stack

- TypeScript
- Office.js (Excel JavaScript API)
- Webpack
- No frameworks — just vanilla TS for speed

## License

MIT

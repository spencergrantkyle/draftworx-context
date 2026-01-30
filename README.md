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
# Clone the repo
git clone https://github.com/spencergrantkyle/draftworx-context.git
cd draftworx-context

# Install dependencies
npm install

# Start dev server (https://localhost:3000)
npm run dev
```

### Windows Setup

1. **Trust the dev certificate** (first time only):
   - When you run `npm run dev`, webpack creates a self-signed HTTPS cert
   - If Excel blocks loading, manually trust the cert or use Edge to visit https://localhost:3000 and accept the warning

2. **Sideload in Excel**:
   - Open Excel
   - Go to: **Insert** → **Add-ins** → **My Add-ins** → **Upload My Add-in**
   - Browse to `manifest.xml` in the project folder
   - Click **Upload**

3. **Use the add-in**:
   - Click the **Draftworx** button in the Home tab
   - Select cells → JSON appears in the taskpane
   - Click **Copy JSON** to clipboard

### Build

```bash
npm run build
```

### Troubleshooting (Windows)

- **"Add-in failed to load"**: Check that `npm run dev` is running and https://localhost:3000 is accessible
- **Certificate errors**: Visit https://localhost:3000 in Edge/Chrome, accept the security warning
- **Taskpane blank**: Check browser console (F12 in taskpane) for errors

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

# Draftworx Context & Automation Library

Excel add-in for AI-powered assistance and composable automations.

## Features

### Context Extraction
- 🔴 **Live Updates** — Automatically captures selection changes
- 📋 **One-Click Copy** — JSON to clipboard instantly
- 📊 **Rich Data** — Cell references, values, and formulas

### Automation Library
- ⚡ **Composable** — Small, reusable functions that chain together
- 🔧 **Primitives** — Low-level building blocks for custom workflows
- 🚀 **Automations** — Pre-built, ready-to-use composed operations

## Architecture

```
src/
├── taskpane/           # Add-in UI
│   ├── taskpane.html
│   └── taskpane.ts
└── lib/                # Automation library
    ├── index.ts        # Public exports
    ├── types.ts        # Type definitions
    ├── range.ts        # Range capture/paste primitives
    ├── sheet.ts        # Sheet management primitives
    └── automations.ts  # Composed automations
```

### Design Principles

1. **Primitives vs Automations**
   - Primitives (`range.ts`, `sheet.ts`) require `Excel.run()` context
   - Automations (`automations.ts`) handle their own context — call directly

2. **Composability**
   - Each function does one thing well
   - Chain primitives for custom workflows
   - Use automations for common patterns

3. **Type Safety**
   - Full TypeScript with strict types
   - Results wrapped in `AutomationResult<T>` for error handling

## Using the Library

### Quick Start — Pre-built Automations

```typescript
import { copySelectionToNewSheet } from './lib';

// Copy selection to a new sheet (values + formulas)
const result = await copySelectionToNewSheet();
if (result.success) {
  console.log(`Created: ${result.data.newSheetName}`);
}

// Values only (strip formulas)
await copySelectionToNewSheet({ valuesOnly: true });
```

### Building Custom Automations

```typescript
import { captureSelection, pasteRange, createSheet } from './lib';

// Create a custom automation that copies selection to multiple sheets
await Excel.run(async (context) => {
  const captured = await captureSelection(context);
  
  // Create 3 backup sheets
  for (let i = 1; i <= 3; i++) {
    const sheet = await createSheet(context, { name: `Backup ${i}` });
    await pasteRange(context, captured, sheet);
  }
});
```

## Available Functions

### Primitives (use inside Excel.run)

| Function | Description |
|----------|-------------|
| `captureSelection(context)` | Capture current selection as CapturedRange |
| `captureRange(context, address, sheet?)` | Capture specific range |
| `pasteRange(context, captured, sheet, address?, valuesOnly?)` | Paste captured data |
| `createSheet(context, options?)` | Create a new worksheet |
| `generateUniqueSheetName(context, baseName)` | Get available sheet name |
| `getActiveSheet(context)` | Get the active worksheet |
| `activateSheet(sheet, context)` | Switch to a worksheet |

### Automations (call directly)

| Function | Description |
|----------|-------------|
| `copySelectionToNewSheet(options?)` | Copy selection to a new sheet |
| `duplicateSelection(targetAddress, valuesOnly?)` | Duplicate within same sheet |

## Development

### Prerequisites

- Node.js 18+ (LTS recommended)
- Excel (Windows, Mac, or Web)

### Setup

```bash
git clone https://github.com/spencergrantkyle/draftworx-context.git
cd draftworx-context
npm install
npm run dev
```

### Build

```bash
npm run build
```

### Sideloading

1. Run `npm run dev`
2. In Excel: **Insert** → **Add-ins** → **My Add-ins** → **Upload My Add-in**
3. Select `manifest.xml`

## JSON Output Format (Context Feature)

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

## Tech Stack

- TypeScript (strict mode)
- Office.js (Excel JavaScript API)
- Webpack

## License

MIT

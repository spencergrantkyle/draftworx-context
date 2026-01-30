/**
 * Draftworx Automation Library
 * 
 * A composable library of Excel automation primitives and higher-level automations.
 * 
 * Architecture:
 * - types.ts     → Type definitions
 * - range.ts     → Range capture/paste primitives (use inside Excel.run)
 * - sheet.ts     → Sheet management primitives (use inside Excel.run)
 * - automations.ts → Composed automations (call directly, handle their own Excel.run)
 * 
 * Usage:
 * - For simple tasks, use the composed automations directly
 * - For custom workflows, import primitives and compose inside your own Excel.run
 * 
 * @example
 * // Using a composed automation
 * import { copySelectionToNewSheet } from './lib';
 * const result = await copySelectionToNewSheet({ valuesOnly: true });
 * 
 * @example
 * // Building a custom automation
 * import { captureSelection, pasteRange, createSheet } from './lib';
 * await Excel.run(async (context) => {
 *   const captured = await captureSelection(context);
 *   const sheet1 = await createSheet(context, { name: 'Backup 1' });
 *   const sheet2 = await createSheet(context, { name: 'Backup 2' });
 *   await pasteRange(context, captured, sheet1);
 *   await pasteRange(context, captured, sheet2, 'A1', true); // values only
 * });
 */

// Types
export * from './types';

// Primitives
export { captureSelection, captureRange, pasteRange } from './range';
export { createSheet, generateUniqueSheetName, getActiveSheet, activateSheet } from './sheet';

// Composed Automations
export { 
  copySelectionToNewSheet, 
  duplicateSelection,
  type CopyToNewSheetOptions,
  type CopyToNewSheetResult 
} from './automations';

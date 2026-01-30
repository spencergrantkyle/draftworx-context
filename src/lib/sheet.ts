/**
 * Draftworx Automation Library - Sheet Utilities
 * 
 * Functions for creating and manipulating Excel worksheets.
 * All functions are designed to be called within an Excel.run() context.
 */

import { CreateSheetOptions } from './types';

/**
 * Create a new worksheet.
 * Must be called within Excel.run().
 * 
 * @param context - Excel RequestContext
 * @param options - Sheet creation options
 * @returns Promise<Excel.Worksheet> - The newly created worksheet
 * 
 * @example
 * await Excel.run(async (context) => {
 *   const newSheet = await createSheet(context, { name: 'Copy of Data' });
 *   newSheet.activate();
 *   await context.sync();
 * });
 */
export async function createSheet(
  context: Excel.RequestContext,
  options: CreateSheetOptions = {}
): Promise<Excel.Worksheet> {
  const sheets = context.workbook.worksheets;
  const activeSheet = sheets.getActiveWorksheet();
  
  activeSheet.load('position');
  await context.sync();
  
  // Add the new sheet
  const newSheet = sheets.add(options.name);
  
  // Position it appropriately
  if (options.position === 'before') {
    newSheet.position = activeSheet.position;
  } else if (options.position === 'end') {
    // Let Excel handle it — default is at the end
  } else {
    // Default: after active sheet
    newSheet.position = activeSheet.position + 1;
  }
  
  await context.sync();
  
  return newSheet;
}

/**
 * Generate a unique sheet name based on a base name.
 * Appends numbers if the name already exists.
 * Must be called within Excel.run().
 * 
 * @param context - Excel RequestContext
 * @param baseName - Base name for the sheet
 * @returns Promise<string> - A unique sheet name
 */
export async function generateUniqueSheetName(
  context: Excel.RequestContext,
  baseName: string
): Promise<string> {
  const sheets = context.workbook.worksheets;
  sheets.load('items/name');
  await context.sync();
  
  const existingNames = new Set(sheets.items.map(s => s.name.toLowerCase()));
  
  // Try the base name first
  if (!existingNames.has(baseName.toLowerCase())) {
    return baseName;
  }
  
  // Append numbers until we find a unique name
  let counter = 2;
  while (existingNames.has(`${baseName} (${counter})`.toLowerCase())) {
    counter++;
  }
  
  return `${baseName} (${counter})`;
}

/**
 * Get the active worksheet.
 * Must be called within Excel.run().
 * 
 * @param context - Excel RequestContext
 * @returns Promise<Excel.Worksheet>
 */
export async function getActiveSheet(context: Excel.RequestContext): Promise<Excel.Worksheet> {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.load(['name', 'position']);
  await context.sync();
  return sheet;
}

/**
 * Activate (switch to) a worksheet.
 * Must be called within Excel.run().
 * 
 * @param sheet - The worksheet to activate
 * @param context - Excel RequestContext
 */
export async function activateSheet(
  sheet: Excel.Worksheet,
  context: Excel.RequestContext
): Promise<void> {
  sheet.activate();
  await context.sync();
}

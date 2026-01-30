/**
 * Draftworx Automation Library - Range Utilities
 * 
 * Functions for capturing and manipulating Excel ranges.
 * All functions are designed to be called within an Excel.run() context.
 */

import { CapturedRange } from './types';

/**
 * Capture the current selection as a CapturedRange object.
 * Must be called within Excel.run().
 * 
 * @param context - Excel RequestContext from Excel.run()
 * @returns Promise<CapturedRange> - The captured range data
 * 
 * @example
 * await Excel.run(async (context) => {
 *   const captured = await captureSelection(context);
 *   console.log(captured.address, captured.values);
 * });
 */
export async function captureSelection(context: Excel.RequestContext): Promise<CapturedRange> {
  const range = context.workbook.getSelectedRange();
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  range.load(['address', 'values', 'formulas', 'rowCount', 'columnCount']);
  sheet.load('name');
  
  await context.sync();
  
  // Extract just the cell reference (remove sheet name prefix)
  const addressOnly = range.address.includes('!')
    ? range.address.split('!')[1]
    : range.address;
  
  return {
    address: addressOnly,
    sourceSheet: sheet.name,
    values: range.values,
    formulas: range.formulas,
    rowCount: range.rowCount,
    columnCount: range.columnCount,
  };
}

/**
 * Capture a specific range by address.
 * Must be called within Excel.run().
 * 
 * @param context - Excel RequestContext
 * @param address - Cell address (e.g., "A1:C10")
 * @param sheetName - Optional sheet name (defaults to active sheet)
 * @returns Promise<CapturedRange>
 */
export async function captureRange(
  context: Excel.RequestContext,
  address: string,
  sheetName?: string
): Promise<CapturedRange> {
  const sheet = sheetName
    ? context.workbook.worksheets.getItem(sheetName)
    : context.workbook.worksheets.getActiveWorksheet();
  
  const range = sheet.getRange(address);
  
  range.load(['address', 'values', 'formulas', 'rowCount', 'columnCount']);
  sheet.load('name');
  
  await context.sync();
  
  const addressOnly = range.address.includes('!')
    ? range.address.split('!')[1]
    : range.address;
  
  return {
    address: addressOnly,
    sourceSheet: sheet.name,
    values: range.values,
    formulas: range.formulas,
    rowCount: range.rowCount,
    columnCount: range.columnCount,
  };
}

/**
 * Paste captured range data to a target location.
 * Must be called within Excel.run().
 * 
 * @param context - Excel RequestContext
 * @param captured - Previously captured range data
 * @param targetSheet - Target worksheet
 * @param targetAddress - Address to paste to (defaults to captured.address)
 * @param valuesOnly - If true, paste values only (no formulas)
 */
export async function pasteRange(
  context: Excel.RequestContext,
  captured: CapturedRange,
  targetSheet: Excel.Worksheet,
  targetAddress?: string,
  valuesOnly: boolean = false
): Promise<void> {
  const address = targetAddress || captured.address;
  const range = targetSheet.getRange(address);
  
  if (valuesOnly) {
    range.values = captured.values;
  } else {
    // Paste formulas — Excel will paste values where there's no formula
    range.formulas = captured.formulas;
  }
  
  await context.sync();
}

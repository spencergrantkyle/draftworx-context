/**
 * Draftworx Automation Library - Composed Automations
 * 
 * Higher-level automations that chain together primitive operations.
 * Each automation is a self-contained, reusable unit.
 */

import { AutomationResult, CapturedRange, CreateSheetOptions, PasteOptions } from './types';
import { captureSelection, pasteRange } from './range';
import { createSheet, generateUniqueSheetName } from './sheet';

/**
 * Result of copySelectionToNewSheet automation
 */
export interface CopyToNewSheetResult {
  /** The captured range data */
  captured: CapturedRange;
  /** Name of the newly created sheet */
  newSheetName: string;
  /** Address where data was pasted */
  pastedAddress: string;
}

/**
 * Options for copySelectionToNewSheet automation
 */
export interface CopyToNewSheetOptions {
  /** Base name for the new sheet (default: "Copy of [source sheet]") */
  sheetName?: string;
  /** Paste values only, stripping formulas */
  valuesOnly?: boolean;
  /** Activate the new sheet after creation (default: true) */
  activateNewSheet?: boolean;
}

/**
 * Copy the current selection to a new sheet.
 * 
 * Creates a new worksheet and pastes the selected range in the same
 * cell location on the new sheet.
 * 
 * This is a complete automation — call it directly, not inside Excel.run().
 * 
 * @param options - Configuration options
 * @returns Promise<AutomationResult<CopyToNewSheetResult>>
 * 
 * @example
 * // Basic usage
 * const result = await copySelectionToNewSheet();
 * if (result.success) {
 *   console.log(`Created ${result.data.newSheetName}`);
 * }
 * 
 * @example
 * // With options
 * const result = await copySelectionToNewSheet({
 *   sheetName: 'Backup',
 *   valuesOnly: true,
 *   activateNewSheet: false
 * });
 */
export async function copySelectionToNewSheet(
  options: CopyToNewSheetOptions = {}
): Promise<AutomationResult<CopyToNewSheetResult>> {
  try {
    return await Excel.run(async (context) => {
      // Step 1: Capture the current selection
      const captured = await captureSelection(context);
      
      if (captured.rowCount === 0 || captured.columnCount === 0) {
        return {
          success: false,
          error: 'No cells selected'
        };
      }
      
      // Step 2: Generate a unique name for the new sheet
      const baseName = options.sheetName || `Copy of ${captured.sourceSheet}`;
      const uniqueName = await generateUniqueSheetName(context, baseName);
      
      // Step 3: Create the new sheet
      const newSheet = await createSheet(context, { name: uniqueName });
      
      // Step 4: Paste the captured range to the same location
      await pasteRange(
        context,
        captured,
        newSheet,
        captured.address,  // Same location
        options.valuesOnly ?? false
      );
      
      // Step 5: Optionally activate the new sheet
      if (options.activateNewSheet !== false) {
        newSheet.activate();
        await context.sync();
      }
      
      return {
        success: true,
        data: {
          captured,
          newSheetName: uniqueName,
          pastedAddress: captured.address
        }
      };
    });
  } catch (error) {
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error)
    };
  }
}

/**
 * Duplicate the current selection within the same sheet.
 * 
 * @param targetAddress - Address to paste to
 * @param valuesOnly - Strip formulas if true
 */
export async function duplicateSelection(
  targetAddress: string,
  valuesOnly: boolean = false
): Promise<AutomationResult<{ sourceAddress: string; targetAddress: string }>> {
  try {
    return await Excel.run(async (context) => {
      const captured = await captureSelection(context);
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      await pasteRange(context, captured, sheet, targetAddress, valuesOnly);
      
      return {
        success: true,
        data: {
          sourceAddress: captured.address,
          targetAddress
        }
      };
    });
  } catch (error) {
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error)
    };
  }
}

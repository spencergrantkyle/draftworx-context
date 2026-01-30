/**
 * Draftworx Automation Library - Type Definitions
 * 
 * Core types used across all automation functions.
 * Keep this lean — add types as needed, not speculatively.
 */

/**
 * Result of a range capture operation
 */
export interface CapturedRange {
  /** Original address (e.g., "A1:C10") */
  address: string;
  /** Sheet name where the range was captured */
  sourceSheet: string;
  /** 2D array of cell values */
  values: (string | number | boolean | null)[][];
  /** 2D array of formulas (empty string if no formula) */
  formulas: string[][];
  /** Number of rows */
  rowCount: number;
  /** Number of columns */
  columnCount: number;
}

/**
 * Options for creating a new sheet
 */
export interface CreateSheetOptions {
  /** Name for the new sheet (auto-generated if not provided) */
  name?: string;
  /** Position to insert the sheet (default: after active sheet) */
  position?: 'before' | 'after' | 'end';
}

/**
 * Options for pasting a range
 */
export interface PasteOptions {
  /** Target address to paste to (default: same as source) */
  targetAddress?: string;
  /** Whether to paste values only (no formulas) */
  valuesOnly?: boolean;
}

/**
 * Result of an automation operation
 */
export interface AutomationResult<T = void> {
  success: boolean;
  data?: T;
  error?: string;
}

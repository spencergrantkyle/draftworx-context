/*
 * Draftworx Context - Excel Add-in
 * Extracts selected cell data as structured JSON for AI assistance
 */

interface CellData {
  ref: string;
  value: string | number | boolean | null;
  formula: string | null;
}

interface ContextPayload {
  selection: string;
  sheet: string;
  timestamp: string;
  cells: CellData[];
}

// State
let currentContext: ContextPayload | null = null;
let selectionHandler: OfficeExtension.EventHandlerResult<Excel.WorksheetSelectionChangedEventArgs> | null = null;

// DOM Elements
const selectionAddressEl = document.getElementById('selectionAddress') as HTMLDivElement;
const cellCountEl = document.getElementById('cellCount') as HTMLDivElement;
const contextJsonEl = document.getElementById('contextJson') as HTMLPreElement;
const copyBtn = document.getElementById('copyBtn') as HTMLButtonElement;
const refreshBtn = document.getElementById('refreshBtn') as HTMLButtonElement;
const statusEl = document.getElementById('status') as HTMLDivElement;
const includeFormulasCheckbox = document.getElementById('includeFormulas') as HTMLInputElement;
const liveUpdateCheckbox = document.getElementById('liveUpdate') as HTMLInputElement;

// Initialize Office
Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    console.log('Draftworx Context initialized');
    
    // Set up event listeners
    copyBtn.addEventListener('click', copyToClipboard);
    refreshBtn.addEventListener('click', () => extractSelectionData());
    liveUpdateCheckbox.addEventListener('change', toggleLiveUpdate);
    includeFormulasCheckbox.addEventListener('change', () => extractSelectionData());
    
    // Initial extraction
    await extractSelectionData();
    
    // Set up live updates if enabled
    if (liveUpdateCheckbox.checked) {
      await setupSelectionHandler();
    }
  }
});

/**
 * Extract data from the current selection
 */
async function extractSelectionData(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Load all needed properties
      range.load(['address', 'values', 'formulas', 'rowCount', 'columnCount']);
      sheet.load('name');
      
      await context.sync();
      
      const includeFormulas = includeFormulasCheckbox.checked;
      const cells: CellData[] = [];
      
      // Parse the address to get column/row info
      const fullAddress = range.address; // e.g., "Sheet1!A1:C3"
      const addressPart = fullAddress.includes('!') 
        ? fullAddress.split('!')[1] 
        : fullAddress;
      
      // Get starting cell reference
      const startCell = addressPart.split(':')[0];
      const startCol = startCell.match(/[A-Z]+/)?.[0] || 'A';
      const startRow = parseInt(startCell.match(/\d+/)?.[0] || '1');
      
      // Process each cell
      for (let row = 0; row < range.rowCount; row++) {
        for (let col = 0; col < range.columnCount; col++) {
          const colLetter = getColumnLetter(columnToNumber(startCol) + col);
          const rowNum = startRow + row;
          const cellRef = `${colLetter}${rowNum}`;
          
          const value = range.values[row][col];
          const formula = range.formulas[row][col];
          
          // Only include formula if it's different from the value (i.e., it's an actual formula)
          const hasFormula = typeof formula === 'string' && formula.startsWith('=');
          
          cells.push({
            ref: cellRef,
            value: value,
            formula: includeFormulas && hasFormula ? formula : null
          });
        }
      }
      
      // Build context payload
      currentContext = {
        selection: addressPart,
        sheet: sheet.name,
        timestamp: new Date().toISOString(),
        cells: cells
      };
      
      // Update UI
      updateUI();
    });
  } catch (error) {
    console.error('Error extracting selection:', error);
    showStatus('Error extracting selection data', 'error');
  }
}

/**
 * Update the UI with current context
 */
function updateUI(): void {
  if (!currentContext) return;
  
  // Update selection info
  selectionAddressEl.textContent = `${currentContext.sheet}!${currentContext.selection}`;
  cellCountEl.textContent = `${currentContext.cells.length} cell${currentContext.cells.length !== 1 ? 's' : ''}`;
  
  // Update JSON display with syntax highlighting
  const jsonString = JSON.stringify(currentContext, null, 2);
  contextJsonEl.innerHTML = syntaxHighlight(jsonString);
}

/**
 * Syntax highlight JSON
 */
function syntaxHighlight(json: string): string {
  return json
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/("(\\u[a-zA-Z0-9]{4}|\\[^u]|[^\\"])*"(\s*:)?|\b(true|false|null)\b|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?)/g, 
      (match) => {
        let cls = 'number';
        if (/^"/.test(match)) {
          cls = /:$/.test(match) ? 'key' : 'string';
        } else if (/true|false/.test(match)) {
          cls = 'boolean';
        } else if (/null/.test(match)) {
          cls = 'null';
        }
        return `<span class="${cls}">${match}</span>`;
      }
    );
}

/**
 * Copy context JSON to clipboard
 */
async function copyToClipboard(): Promise<void> {
  if (!currentContext) {
    showStatus('No data to copy', 'error');
    return;
  }
  
  try {
    const jsonString = JSON.stringify(currentContext, null, 2);
    await navigator.clipboard.writeText(jsonString);
    showStatus('Copied to clipboard!', 'success');
  } catch (error) {
    // Fallback for older browsers
    const textarea = document.createElement('textarea');
    textarea.value = JSON.stringify(currentContext, null, 2);
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand('copy');
    document.body.removeChild(textarea);
    showStatus('Copied to clipboard!', 'success');
  }
}

/**
 * Show status message
 */
function showStatus(message: string, type: 'success' | 'error'): void {
  statusEl.textContent = message;
  statusEl.className = `status ${type}`;
  
  setTimeout(() => {
    statusEl.className = 'status';
  }, 3000);
}

/**
 * Toggle live update on selection change
 */
async function toggleLiveUpdate(): Promise<void> {
  if (liveUpdateCheckbox.checked) {
    await setupSelectionHandler();
  } else {
    await removeSelectionHandler();
  }
}

/**
 * Set up selection change handler
 */
async function setupSelectionHandler(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      selectionHandler = sheet.onSelectionChanged.add(async () => {
        await extractSelectionData();
      });
      await context.sync();
      console.log('Selection handler registered');
    });
  } catch (error) {
    console.error('Error setting up selection handler:', error);
  }
}

/**
 * Remove selection change handler
 */
async function removeSelectionHandler(): Promise<void> {
  if (selectionHandler) {
    try {
      await Excel.run(selectionHandler.context, async (context) => {
        selectionHandler!.remove();
        await context.sync();
        selectionHandler = null;
        console.log('Selection handler removed');
      });
    } catch (error) {
      console.error('Error removing selection handler:', error);
    }
  }
}

/**
 * Convert column letter to number (A=1, B=2, ..., Z=26, AA=27, etc.)
 */
function columnToNumber(col: string): number {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64);
  }
  return num;
}

/**
 * Convert column number to letter
 */
function getColumnLetter(num: number): string {
  let letter = '';
  while (num > 0) {
    const mod = (num - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    num = Math.floor((num - 1) / 26);
  }
  return letter;
}

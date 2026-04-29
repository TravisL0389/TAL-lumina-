import { SheetState, CellData } from '../types';

const parseCellId = (id: string) => {
  const colMatch = id.match(/[A-Z]+/);
  const rowMatch = id.match(/[0-9]+/);
  if (!colMatch || !rowMatch) return null;
  const col = colMatch[0];
  const row = parseInt(rowMatch[0]);
  let colIndex = 0;
  for (let i = 0; i < col.length; i++) {
    colIndex = colIndex * 26 + (col.charCodeAt(i) - 64);
  }
  return { r: row, c: colIndex, id: `${col}${row}` };
};

const getCellIdFromCoords = (r: number, c: number) => {
  let col = "";
  let tempC = c;
  while (tempC > 0) {
    let mod = (tempC - 1) % 26;
    col = String.fromCharCode(65 + mod) + col;
    tempC = Math.floor((tempC - mod) / 26);
  }
  return `${col}${r}`;
};

const getRangeCells = (rangeStr: string): string[] => {
  const parts = rangeStr.split(':');
  if (parts.length === 1) return [parts[0].trim()];
  if (parts.length !== 2) return [];
  
  const start = parseCellId(parts[0].trim());
  const end = parseCellId(parts[1].trim());
  
  if (!start || !end) return [];
  
  const r1 = Math.min(start.r, end.r);
  const r2 = Math.max(start.r, end.r);
  const c1 = Math.min(start.c, end.c);
  const c2 = Math.max(start.c, end.c);
  
  const cells = [];
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      cells.push(getCellIdFromCoords(r, c));
    }
  }
  return cells;
};

const getCellValue = (id: string, sheet: SheetState): number | string => {
  const cell = sheet.cells[id];
  if (!cell) return 0;
  const val = cell.displayValue || cell.value || '';
  const num = parseFloat(val);
  return isNaN(num) ? val : num;
};

const handleAggregate = (expr: string, sheet: SheetState, aggFn: (arr: number[]) => number): string => {
  const match = expr.match(/^[A-Z]+\((.*)\)$/);
  if (!match) return '#ERROR!';
  
  const args = match[1].split(',').map(s => s.trim());
  let values: number[] = [];
  
  for (const arg of args) {
    const cells = getRangeCells(arg);
    for (const cellId of cells) {
      const val = getCellValue(cellId, sheet);
      if (typeof val === 'number') {
        values.push(val);
      }
    }
  }
  
  if (values.length === 0 && expr.startsWith('COUNT')) return '0';
  if (values.length === 0) return '0';
  
  return aggFn(values).toString();
};

const handleVlookup = (expr: string, sheet: SheetState): string => {
  // VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
  const match = expr.match(/^VLOOKUP\((.*)\)$/);
  if (!match) return '#ERROR!';
  
  // Split by comma, but ignore commas inside quotes (simple version)
  const args = match[1].split(',').map(s => s.trim());
  if (args.length < 3) return '#N/A';
  
  let lookupValue: string | number = args[0];
  if (lookupValue.startsWith('"') && lookupValue.endsWith('"')) {
    lookupValue = lookupValue.slice(1, -1);
  } else if (!isNaN(parseFloat(lookupValue))) {
    lookupValue = parseFloat(lookupValue);
  } else {
    lookupValue = getCellValue(lookupValue, sheet);
  }
  
  const tableArray = getRangeCells(args[1]);
  const colIndexNum = parseInt(args[2]);
  const exactMatch = args.length > 3 ? args[3].toUpperCase() === 'FALSE' || args[3] === '0' : true;
  
  if (isNaN(colIndexNum) || colIndexNum < 1) return '#VALUE!';
  if (tableArray.length === 0) return '#N/A';
  
  // Determine table dimensions
  const startCell = parseCellId(tableArray[0]);
  const endCell = parseCellId(tableArray[tableArray.length - 1]);
  if (!startCell || !endCell) return '#N/A';
  
  const numCols = endCell.c - startCell.c + 1;
  if (colIndexNum > numCols) return '#REF!';
  
  // Search first column
  for (let r = startCell.r; r <= endCell.r; r++) {
    const searchCellId = getCellIdFromCoords(r, startCell.c);
    const cellVal = getCellValue(searchCellId, sheet);
    
    let isMatch = false;
    if (exactMatch) {
      isMatch = cellVal.toString().toLowerCase() === lookupValue.toString().toLowerCase();
    } else {
      // Approximate match logic (simplified)
      isMatch = cellVal.toString().toLowerCase().includes(lookupValue.toString().toLowerCase());
    }
    
    if (isMatch) {
      const resultCellId = getCellIdFromCoords(r, startCell.c + colIndexNum - 1);
      return getCellValue(resultCellId, sheet).toString();
    }
  }
  
  return '#N/A';
};

const evaluateMath = (expr: string, sheet: SheetState): string => {
  // Very basic math evaluator replacing cell refs with values
  // e.g. A1+B1*2
  let evalStr = expr;
  const cellRefs = expr.match(/[A-Z]+[0-9]+/g) || [];
  
  for (const ref of cellRefs) {
    const val = getCellValue(ref, sheet);
    const numVal = typeof val === 'number' ? val : 0;
    evalStr = evalStr.replace(new RegExp(ref, 'g'), numVal.toString());
  }
  
  try {
    // eslint-disable-next-line no-new-func
    const result = new Function(`return ${evalStr}`)();
    return isNaN(result) ? '#VALUE!' : result.toString();
  } catch (e) {
    return '#ERROR!';
  }
};

export const evaluateFormula = (formula: string, sheet: SheetState): string => {
  if (!formula.startsWith('=')) return formula;
  
  const expr = formula.substring(1).trim().toUpperCase();
  
  try {
    if (expr.startsWith('SUM(')) return handleAggregate(expr, sheet, (arr) => arr.reduce((a, b) => a + b, 0));
    if (expr.startsWith('AVERAGE(')) return handleAggregate(expr, sheet, (arr) => arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0);
    if (expr.startsWith('MIN(')) return handleAggregate(expr, sheet, (arr) => Math.min(...arr));
    if (expr.startsWith('MAX(')) return handleAggregate(expr, sheet, (arr) => Math.max(...arr));
    if (expr.startsWith('COUNT(')) return handleAggregate(expr, sheet, (arr) => arr.length);
    if (expr.startsWith('VLOOKUP(')) return handleVlookup(expr, sheet);
    
    return evaluateMath(expr, sheet);
  } catch (e) {
    return '#ERROR!';
  }
};

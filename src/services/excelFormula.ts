import { CellData, CellFormatType, Language } from '../types/office';

// Convert column index (0-based) to letter: 0 -> A, 25 -> Z, 26 -> AA
export function indexToColName(index: number): string {
  let colName = '';
  let temp = index;
  while (temp >= 0) {
    colName = String.fromCharCode((temp % 26) + 65) + colName;
    temp = Math.floor(temp / 26) - 1;
  }
  return colName;
}

// Convert letter to column index (0-based): 'A' -> 0, 'Z' -> 25, 'AA' -> 26
export function colNameToIndex(colName: string): number {
  const upper = colName.toUpperCase();
  let index = 0;
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index - 1;
}

// Parse cell coordinate like 'A1' into col and row
export function parseCellRef(ref: string): { col: string; colIdx: number; row: number; rowIdx: number } | null {
  const match = ref.trim().toUpperCase().match(/^([A-Z]+)([0-9]+)$/);
  if (!match) return null;
  const col = match[1];
  const row = parseInt(match[2], 10);
  return {
    col,
    colIdx: colNameToIndex(col),
    row,
    rowIdx: row - 1,
  };
}

// Expand a range like 'A1:B3' into an array of cell coordinates ['A1', 'A2', 'A3', 'B1', 'B2', 'B3']
export function expandCellRange(rangeStr: string): string[] {
  const parts = rangeStr.split(':');
  if (parts.length === 1) {
    return [parts[0].trim().toUpperCase()];
  }
  const start = parseCellRef(parts[0]);
  const end = parseCellRef(parts[1]);
  if (!start || !end) return [];

  const minCol = Math.min(start.colIdx, end.colIdx);
  const maxCol = Math.max(start.colIdx, end.colIdx);
  const minRow = Math.min(start.row, end.row);
  const maxRow = Math.max(start.row, end.row);

  const cells: string[] = [];
  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      cells.push(`${indexToColName(c)}${r}`);
    }
  }
  return cells;
}

// Get the numeric value of a cell
export function getNumericValueOfCell(
  cellRef: string,
  cells: Record<string, CellData>,
  visited: Set<string> = new Set()
): number {
  const val = evaluateCell(cellRef, cells, visited);
  if (typeof val === 'number') return val;
  const parsed = parseFloat(String(val));
  return isNaN(parsed) ? 0 : parsed;
}

// Main cell evaluation function
export function evaluateCell(
  cellRef: string,
  cells: Record<string, CellData>,
  visited: Set<string> = new Set()
): string | number {
  const key = cellRef.toUpperCase();
  const cell = cells[key];
  if (!cell || cell.raw === undefined || cell.raw === null || cell.raw === '') {
    return '';
  }

  const raw = String(cell.raw).trim();
  if (!raw.startsWith('=')) {
    // If it's a plain number
    const num = Number(raw);
    if (!isNaN(num) && raw !== '') {
      return num;
    }
    return raw;
  }

  // Circular reference detection
  if (visited.has(key)) {
    return '#REF!';
  }
  visited.add(key);

  try {
    const result = evaluateFormulaString(raw.slice(1).trim(), cells, visited);
    visited.delete(key);
    return result;
  } catch {
    visited.delete(key);
    return '#ERROR!';
  }
}

// Evaluates a formula without the leading '='
export function evaluateFormulaString(
  formula: string,
  cells: Record<string, CellData>,
  visited: Set<string>
): string | number {
  const upper = formula.trim().toUpperCase();

  // 1. Check for Functions: SUM, AVERAGE, COUNT, MIN, MAX, IF, ROUND, CONCAT
  const funcMatch = upper.match(/^([A-Z]+)\((.*)\)$/);
  if (funcMatch) {
    const funcName = funcMatch[1];
    const argsStr = funcMatch[2];

    if (funcName === 'SUM') {
      const cellRefs = parseArgsToCellList(argsStr);
      let sum = 0;
      for (const ref of cellRefs) {
        sum += getNumericValueOfCell(ref, cells, new Set(visited));
      }
      return sum;
    }

    if (funcName === 'AVERAGE') {
      const cellRefs = parseArgsToCellList(argsStr);
      if (cellRefs.length === 0) return 0;
      let sum = 0;
      let count = 0;
      for (const ref of cellRefs) {
        const val = evaluateCell(ref, cells, new Set(visited));
        if (typeof val === 'number' || (!isNaN(Number(val)) && val !== '')) {
          sum += Number(val);
          count++;
        }
      }
      return count === 0 ? 0 : Number((sum / count).toFixed(4));
    }

    if (funcName === 'COUNT') {
      const cellRefs = parseArgsToCellList(argsStr);
      let count = 0;
      for (const ref of cellRefs) {
        const val = evaluateCell(ref, cells, new Set(visited));
        if (typeof val === 'number' || (!isNaN(Number(val)) && val !== '')) {
          count++;
        }
      }
      return count;
    }

    if (funcName === 'MIN') {
      const cellRefs = parseArgsToCellList(argsStr);
      if (cellRefs.length === 0) return 0;
      let minVal = Infinity;
      let found = false;
      for (const ref of cellRefs) {
        const val = evaluateCell(ref, cells, new Set(visited));
        if (typeof val === 'number' || (!isNaN(Number(val)) && val !== '')) {
          minVal = Math.min(minVal, Number(val));
          found = true;
        }
      }
      return found ? minVal : 0;
    }

    if (funcName === 'MAX') {
      const cellRefs = parseArgsToCellList(argsStr);
      if (cellRefs.length === 0) return 0;
      let maxVal = -Infinity;
      let found = false;
      for (const ref of cellRefs) {
        const val = evaluateCell(ref, cells, new Set(visited));
        if (typeof val === 'number' || (!isNaN(Number(val)) && val !== '')) {
          maxVal = Math.max(maxVal, Number(val));
          found = true;
        }
      }
      return found ? maxVal : 0;
    }

    if (funcName === 'IF') {
      // Split args safely by comma not inside quotes
      const args = splitArguments(argsStr);
      if (args.length >= 2) {
        const conditionStr = args[0];
        const trueResult = evaluateExpression(args[1], cells, visited);
        const falseResult = args.length >= 3 ? evaluateExpression(args[2], cells, visited) : '';

        const condVal = evaluateCondition(conditionStr, cells, visited);
        return condVal ? trueResult : falseResult;
      }
      return '#VALUE!';
    }

    if (funcName === 'ROUND') {
      const args = splitArguments(argsStr);
      if (args.length >= 1) {
        const numVal = Number(evaluateExpression(args[0], cells, visited));
        const decimals = args[1] ? Number(evaluateExpression(args[1], cells, visited)) : 0;
        return isNaN(numVal) ? '#VALUE!' : Number(numVal.toFixed(decimals));
      }
      return '#VALUE!';
    }

    if (funcName === 'CONCAT' || funcName === 'CONCATENATE') {
      const args = splitArguments(argsStr);
      let res = '';
      for (const arg of args) {
        res += String(evaluateExpression(arg, cells, visited));
      }
      return res;
    }
  }

  // 2. Fallback: mathematical expression like A1 + B2 * 5
  return evaluateExpression(upper, cells, visited);
}

// Parses comma-separated arguments that might include ranges like A1:A5, B2
function parseArgsToCellList(argsStr: string): string[] {
  const parts = splitArguments(argsStr);
  const result: string[] = [];
  for (const part of parts) {
    const trimmed = part.trim();
    if (trimmed.includes(':')) {
      result.push(...expandCellRange(trimmed));
    } else {
      result.push(trimmed);
    }
  }
  return result;
}

// Split arguments by comma while respecting quotes and parens
function splitArguments(str: string): string[] {
  const result: string[] = [];
  let current = '';
  let inQuotes = false;
  let parenDepth = 0;

  for (let i = 0; i < str.length; i++) {
    const char = str[i];
    if (char === '"' || char === "'") {
      inQuotes = !inQuotes;
      current += char;
    } else if (char === '(' && !inQuotes) {
      parenDepth++;
      current += char;
    } else if (char === ')' && !inQuotes) {
      parenDepth--;
      current += char;
    } else if (char === ',' && !inQuotes && parenDepth === 0) {
      result.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  if (current.trim().length > 0) {
    result.push(current.trim());
  }
  return result;
}

// Evaluate condition for IF like "A1 > 50" or "B2 = 100"
function evaluateCondition(cond: string, cells: Record<string, CellData>, visited: Set<string>): boolean {
  const opMatch = cond.match(/(>=|<=|!=|<>|>|<|=)/);
  if (!opMatch) {
    const val = evaluateExpression(cond, cells, visited);
    return Boolean(val && val !== '0' && val !== 0);
  }
  const op = opMatch[1];
  const idx = cond.indexOf(op);
  const leftStr = cond.substring(0, idx).trim();
  const rightStr = cond.substring(idx + op.length).trim();

  const left = evaluateExpression(leftStr, cells, visited);
  const right = evaluateExpression(rightStr, cells, visited);

  const numLeft = Number(left);
  const numRight = Number(right);
  const isNumeric = !isNaN(numLeft) && !isNaN(numRight);

  switch (op) {
    case '>':
      return isNumeric ? numLeft > numRight : String(left) > String(right);
    case '<':
      return isNumeric ? numLeft < numRight : String(left) < String(right);
    case '>=':
      return isNumeric ? numLeft >= numRight : String(left) >= String(right);
    case '<=':
      return isNumeric ? numLeft <= numRight : String(left) <= String(right);
    case '=':
      return isNumeric ? numLeft === numRight : String(left).toLowerCase() === String(right).toLowerCase();
    case '!=':
    case '<>':
      return isNumeric ? numLeft !== numRight : String(left).toLowerCase() !== String(right).toLowerCase();
    default:
      return false;
  }
}

// Safe expression evaluator replacing cell references like A1, B2 with numbers
export function evaluateExpression(
  expr: string,
  cells: Record<string, CellData>,
  visited: Set<string>
): string | number {
  let clean = expr.trim();

  // If wrapped in quotes, it's a string literal
  if ((clean.startsWith('"') && clean.endsWith('"')) || (clean.startsWith("'") && clean.endsWith("'"))) {
    return clean.slice(1, -1);
  }

  // Replace cell references with their computed values
  // Regex looks for words like A1, B12, AA5
  const replaced = clean.replace(/\b([A-Z]+[0-9]+)\b/g, (match) => {
    const val = evaluateCell(match, cells, new Set(visited));
    if (typeof val === 'number') return String(val);
    if (typeof val === 'string') {
      const num = Number(val);
      if (!isNaN(num) && val.trim() !== '') return String(num);
      return `"${val.replace(/"/g, '\\"')}"`;
    }
    return '0';
  });

  // Evaluate simple math safely without arbitrary eval
  try {
    // Only allow digits, operators, parens, decimal points, spaces
    if (/^[0-9+\-*/().\s^%]+$/.test(replaced)) {
      // replace ^ with **
      const powerAdjusted = replaced.replace(/\^/g, '**');
      // eslint-disable-next-line no-new-func
      const result = Function(`"use strict"; return (${powerAdjusted});`)();
      if (typeof result === 'number' && !isNaN(result)) {
        return Math.round(result * 10000) / 10000;
      }
      return result;
    }
  } catch {
    // return raw
  }

  return clean;
}

// Format cell value for display
export function formatCellValue(
  value: string | number,
  format?: CellFormatType,
  decimals: number = 2,
  lang: Language = 'en'
): string {
  if (value === '' || value === null || value === undefined) return '';
  if (typeof value === 'string' && (value.startsWith('#') || isNaN(Number(value)))) {
    return value;
  }

  const num = Number(value);
  if (isNaN(num)) return String(value);

  switch (format) {
    case 'currency':
      return lang === 'ar'
        ? `${num.toLocaleString('ar-SA', { minimumFractionDigits: decimals, maximumFractionDigits: decimals })} ر.س`
        : `$${num.toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals })}`;
    case 'percent':
      return `${(num * (num <= 1 && num >= -1 ? 100 : 1)).toFixed(decimals)}%`;
    case 'number':
      return num.toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
    default:
      return num.toLocaleString('en-US', { maximumFractionDigits: 6 });
  }
}

// Calculate summary statistics for selected cells
export interface RangeStats {
  sum: number;
  average: number;
  count: number;
  min: number;
  max: number;
  numericCount: number;
}

export function calculateRangeStats(cellRefs: string[], cells: Record<string, CellData>): RangeStats {
  let sum = 0;
  let count = cellRefs.length;
  let numericCount = 0;
  let min = Infinity;
  let max = -Infinity;

  for (const ref of cellRefs) {
    const val = evaluateCell(ref, cells);
    if (typeof val === 'number' || (!isNaN(Number(val)) && String(val).trim() !== '')) {
      const n = Number(val);
      sum += n;
      numericCount++;
      if (n < min) min = n;
      if (n > max) max = n;
    }
  }

  return {
    sum: Math.round(sum * 100) / 100,
    average: numericCount > 0 ? Math.round((sum / numericCount) * 100) / 100 : 0,
    count,
    min: numericCount > 0 ? min : 0,
    max: numericCount > 0 ? max : 0,
    numericCount,
  };
}

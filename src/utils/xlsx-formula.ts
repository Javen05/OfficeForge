function columnIndexFromLabel(label: string) {
  let result = 0;
  for (let i = 0; i < label.length; i += 1) {
    result = result * 26 + (label.charCodeAt(i) - 64);
  }
  return result - 1;
}

function getCellNumericValue(rows: string[][], ref: string) {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) return 0;
  const col = columnIndexFromLabel(match[1]);
  const row = Number(match[2]) - 1;
  const value = rows[row]?.[col] ?? '';
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function sumRange(rows: string[][], range: string) {
  const [start, end] = range.split(':');
  const startMatch = start.match(/^([A-Z]+)(\d+)$/);
  const endMatch = end.match(/^([A-Z]+)(\d+)$/);
  if (!startMatch || !endMatch) return 0;

  const startCol = columnIndexFromLabel(startMatch[1]);
  const endCol = columnIndexFromLabel(endMatch[1]);
  const startRow = Number(startMatch[2]) - 1;
  const endRow = Number(endMatch[2]) - 1;

  let total = 0;
  for (let row = Math.min(startRow, endRow); row <= Math.max(startRow, endRow); row += 1) {
    for (let col = Math.min(startCol, endCol); col <= Math.max(startCol, endCol); col += 1) {
      const value = rows[row]?.[col] ?? '';
      const parsed = Number.parseFloat(value);
      if (Number.isFinite(parsed)) {
        total += parsed;
      }
    }
  }
  return total;
}

export function evaluateFormula(input: string, rows: string[][]) {
  if (!input.trim().startsWith('=')) return input;

  let expression = input.trim().slice(1).toUpperCase();

  expression = expression.replace(/SUM\(([A-Z]+\d+:[A-Z]+\d+)\)/g, (_, range: string) => `${sumRange(rows, range)}`);
  expression = expression.replace(/([A-Z]+\d+)/g, (ref: string) => `${getCellNumericValue(rows, ref)}`);

  if (!/^[0-9+\-*/().\s]+$/.test(expression)) {
    return input;
  }

  try {
    const result = Function(`"use strict"; return (${expression});`)();
    if (result === null || result === undefined || Number.isNaN(result)) {
      return input;
    }
    return `${result}`;
  } catch {
    return input;
  }
}

'use client';

import { useEffect, useMemo, useRef, useState } from 'react';
import { columnLabel } from '@/utils/document-utils';

export type CellAddress = { row: number; col: number };
export type CellRange = { start: CellAddress; end: CellAddress };

type XlsxPaneProps = {
  rows: string[][];
  formulas: Record<string, string>;
  cellStyles: Record<string, { bold?: boolean; italic?: boolean; highlight?: boolean }>;
  rowHeights: Record<number, number>;
  colWidths: Record<number, number>;
  selectedCell: { row: number; col: number } | null;
  locked: boolean;
  onSelectCell: (cell: { row: number; col: number } | null) => void;
  onCellChange: (row: number, col: number, value: string) => void;
  onApplyFormulaToRange: (value: string, range: CellRange) => void | Promise<void>;
  onToggleBold: (range: CellRange) => void;
  onToggleItalic: (range: CellRange) => void;
  onToggleHighlight: (range: CellRange) => void;
  onAddRow: () => void;
  onAddColumn: () => void;
  onResizeRow: (row: number, height: number) => void;
  onResizeColumn: (col: number, width: number) => void;
};

function cellKey(row: number, col: number) {
  return `${row}:${col}`;
}

function normalizeRange(range: CellRange): CellRange {
  const top = Math.min(range.start.row, range.end.row);
  const bottom = Math.max(range.start.row, range.end.row);
  const left = Math.min(range.start.col, range.end.col);
  const right = Math.max(range.start.col, range.end.col);
  return { start: { row: top, col: left }, end: { row: bottom, col: right } };
}

function isInRange(cell: CellAddress, range: CellRange) {
  const normalized = normalizeRange(range);
  return (
    cell.row >= normalized.start.row &&
    cell.row <= normalized.end.row &&
    cell.col >= normalized.start.col &&
    cell.col <= normalized.end.col
  );
}

function rangeLabel(range: CellRange | null) {
  if (!range) return 'Cell';
  const normalized = normalizeRange(range);
  const start = `${columnLabel(normalized.start.col)}${normalized.start.row + 1}`;
  const end = `${columnLabel(normalized.end.col)}${normalized.end.row + 1}`;
  return start === end ? start : `${start}:${end}`;
}

const DEFAULT_ROW_HEIGHT = 56;
const DEFAULT_COL_WIDTH = 180;

function estimateAutoRowHeight(value: string, columnWidth: number) {
  const text = value || '';
  const lines = text.split('\n');
  const charsPerLine = Math.max(8, Math.floor((columnWidth - 24) / 8));
  const wrappedLineCount = lines.reduce((count, line) => count + Math.max(1, Math.ceil(line.length / charsPerLine)), 0);
  return Math.max(DEFAULT_ROW_HEIGHT, wrappedLineCount * 22 + 16);
}

export function XlsxPane({
  rows,
  formulas,
  cellStyles,
  rowHeights,
  colWidths,
  selectedCell,
  locked,
  onSelectCell,
  onCellChange,
  onApplyFormulaToRange,
  onToggleBold,
  onToggleItalic,
  onToggleHighlight,
  onAddRow,
  onAddColumn,
  onResizeRow,
  onResizeColumn
}: XlsxPaneProps) {
  const maxColumns = rows.length > 0 ? Math.max(...rows.map((row) => row.length), 1) : 1;
  const activeCellValue = selectedCell ? (formulas[cellKey(selectedCell.row, selectedCell.col)] ?? rows[selectedCell.row]?.[selectedCell.col] ?? '') : '';
  const [formulaInput, setFormulaInput] = useState(activeCellValue);
  const [dragging, setDragging] = useState(false);
  const [anchor, setAnchor] = useState<CellAddress | null>(selectedCell);
  const [edge, setEdge] = useState<CellAddress | null>(selectedCell);
  const [stylePopupOpen, setStylePopupOpen] = useState(false);
  const [rowResizeState, setRowResizeState] = useState<{ row: number; startY: number; startHeight: number } | null>(null);
  const [colResizeState, setColResizeState] = useState<{ col: number; startX: number; startWidth: number } | null>(null);
  const stylePopupRef = useRef<HTMLDivElement | null>(null);
  const styleButtonRef = useRef<HTMLButtonElement | null>(null);

  useEffect(() => {
    setFormulaInput(activeCellValue);
  }, [activeCellValue]);

  useEffect(() => {
    const handleMouseUp = () => setDragging(false);
    window.addEventListener('mouseup', handleMouseUp);
    return () => window.removeEventListener('mouseup', handleMouseUp);
  }, []);

  useEffect(() => {
    if (!rowResizeState && !colResizeState) return;

    const handlePointerMove = (event: PointerEvent) => {
      if (rowResizeState) {
        const nextHeight = Math.max(36, rowResizeState.startHeight + (event.clientY - rowResizeState.startY));
        onResizeRow(rowResizeState.row, Math.round(nextHeight));
      }
      if (colResizeState) {
        const nextWidth = Math.max(90, colResizeState.startWidth + (event.clientX - colResizeState.startX));
        onResizeColumn(colResizeState.col, Math.round(nextWidth));
      }
    };

    const handlePointerUp = () => {
      setRowResizeState(null);
      setColResizeState(null);
    };

    window.addEventListener('pointermove', handlePointerMove);
    window.addEventListener('pointerup', handlePointerUp);

    return () => {
      window.removeEventListener('pointermove', handlePointerMove);
      window.removeEventListener('pointerup', handlePointerUp);
    };
  }, [rowResizeState, colResizeState, onResizeColumn, onResizeRow]);

  useEffect(() => {
    if (!stylePopupOpen) return;

    const handlePointerDown = (event: MouseEvent) => {
      const target = event.target as Node;
      if (stylePopupRef.current?.contains(target)) return;
      if (styleButtonRef.current?.contains(target)) return;
      setStylePopupOpen(false);
    };

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setStylePopupOpen(false);
      }
    };

    window.addEventListener('mousedown', handlePointerDown);
    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('mousedown', handlePointerDown);
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [stylePopupOpen]);

  const activeRange = useMemo(() => {
    if (anchor && edge) return normalizeRange({ start: anchor, end: edge });
    if (selectedCell) return { start: selectedCell, end: selectedCell };
    return null;
  }, [anchor, edge, selectedCell]);

  useEffect(() => {
    if (!activeRange) {
      setStylePopupOpen(false);
    }
  }, [activeRange]);

  const allCellsBold = useMemo(() => {
    if (!activeRange) return false;
    for (let row = activeRange.start.row; row <= activeRange.end.row; row += 1) {
      for (let col = activeRange.start.col; col <= activeRange.end.col; col += 1) {
        if (!cellStyles[cellKey(row, col)]?.bold) {
          return false;
        }
      }
    }
    return true;
  }, [activeRange, cellStyles]);

  const allCellsItalic = useMemo(() => {
    if (!activeRange) return false;
    for (let row = activeRange.start.row; row <= activeRange.end.row; row += 1) {
      for (let col = activeRange.start.col; col <= activeRange.end.col; col += 1) {
        if (!cellStyles[cellKey(row, col)]?.italic) {
          return false;
        }
      }
    }
    return true;
  }, [activeRange, cellStyles]);

  const allCellsHighlighted = useMemo(() => {
    if (!activeRange) return false;
    for (let row = activeRange.start.row; row <= activeRange.end.row; row += 1) {
      for (let col = activeRange.start.col; col <= activeRange.end.col; col += 1) {
        if (!cellStyles[cellKey(row, col)]?.highlight) {
          return false;
        }
      }
    }
    return true;
  }, [activeRange, cellStyles]);

  const applyFromFormulaBar = () => {
    if (!activeRange || locked) return;
    onApplyFormulaToRange(formulaInput, activeRange);
  };

  const openStylePopup = () => {
    if (!activeRange || locked) return;
    setStylePopupOpen((current) => !current);
  };

  return (
    <div className="space-y-3 rounded-[24px] border border-white/10 bg-[#07111f] p-4 sm:p-5">
      <div className="sticky top-0 z-40 flex items-center gap-3 rounded-[16px] border border-white/10 bg-white/5 px-3 py-2 shadow-soft">
        <div className="text-xs font-semibold uppercase tracking-[0.2em] text-white/45">
          {rangeLabel(activeRange)}
        </div>
        <input
          value={formulaInput}
          onChange={(event) => setFormulaInput(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              applyFromFormulaBar();
            }
          }}
          placeholder="Formula / value"
          disabled={locked}
          className="w-full bg-transparent text-sm text-white outline-none placeholder:text-white/35"
        />
        <button
          type="button"
          onClick={applyFromFormulaBar}
          disabled={!activeRange || locked}
          className="rounded-full border border-white/10 bg-white/5 px-3 py-1.5 text-xs text-white transition enabled:hover:bg-white/10 disabled:cursor-not-allowed disabled:opacity-40"
          title="Apply current formula/value to all selected cells"
        >
          Apply
        </button>
        <button
          ref={styleButtonRef}
          type="button"
          onClick={openStylePopup}
          disabled={!activeRange || locked}
          className={`rounded-full border px-3 py-1.5 text-xs transition disabled:cursor-not-allowed disabled:opacity-40 ${stylePopupOpen || allCellsBold || allCellsItalic || allCellsHighlighted ? 'border-[#6d7dff]/60 bg-[#6d7dff]/20 text-white' : 'border-white/10 bg-white/5 text-white enabled:hover:bg-white/10'}`}
        >
          Style
        </button>
        <button
          type="button"
          onClick={onAddRow}
          disabled={locked}
          className="rounded-full border border-white/10 bg-white/5 px-3 py-1.5 text-xs text-white transition enabled:hover:bg-white/10 disabled:cursor-not-allowed disabled:opacity-40"
        >
          Add row
        </button>
        <button
          type="button"
          onClick={onAddColumn}
          disabled={locked}
          className="rounded-full border border-white/10 bg-white/5 px-3 py-1.5 text-xs text-white transition enabled:hover:bg-white/10 disabled:cursor-not-allowed disabled:opacity-40"
        >
          Add column
        </button>
      </div>
      <div className="text-xs text-white/55">Apply is for bulk fill. Normal cell edits already auto-update dependent formulas.</div>
      {locked && <div className="text-xs text-[#f6c76a]">Calculating formulas: cells are temporarily locked.</div>}
      {stylePopupOpen && activeRange && !locked && (
        <div ref={stylePopupRef} className="absolute left-4 top-[3.7rem] z-50 w-52 rounded-2xl border border-white/10 bg-[#07111f] p-2 shadow-soft">
          <div className="mb-2 px-2 text-[10px] uppercase tracking-[0.16em] text-white/40">Text style</div>
          <div className="flex flex-col gap-1">
            <button type="button" onClick={() => onToggleBold(activeRange)} className={`rounded-xl px-3 py-2 text-left text-sm transition ${allCellsBold ? 'bg-[#6d7dff]/20 text-white' : 'bg-white/5 text-white/85 hover:bg-white/10'}`}>Bold</button>
            <button type="button" onClick={() => onToggleItalic(activeRange)} className={`rounded-xl px-3 py-2 text-left text-sm transition ${allCellsItalic ? 'bg-[#6d7dff]/20 text-white' : 'bg-white/5 text-white/85 hover:bg-white/10'}`}>Italic</button>
            <button type="button" onClick={() => onToggleHighlight(activeRange)} className={`rounded-xl px-3 py-2 text-left text-sm transition ${allCellsHighlighted ? 'bg-[#6d7dff]/20 text-white' : 'bg-white/5 text-white/85 hover:bg-white/10'}`}>Highlight</button>
          </div>
        </div>
      )}
      <div className="overflow-x-auto overflow-y-auto rounded-[22px] border border-white/10">
        <table className="w-max min-w-full border-collapse bg-[#07111f] text-sm">
          <thead className="sticky top-0 z-10 bg-[#0a162b]">
            <tr>
              <th className="w-14 border-b border-r border-white/10 px-3 py-2 text-left text-xs text-white/45">#</th>
              {Array.from({ length: maxColumns }).map((_, index) => (
                <th key={index} className="relative border-b border-r border-white/10 px-3 py-2 text-left text-xs font-semibold text-white/65 last:border-r-0" style={{ width: `${colWidths[index] ?? DEFAULT_COL_WIDTH}px`, minWidth: `${colWidths[index] ?? DEFAULT_COL_WIDTH}px` }}>
                  <span>{columnLabel(index)}</span>
                  <button
                    type="button"
                    aria-label={`Resize column ${columnLabel(index)}`}
                    onPointerDown={(event) => {
                      event.preventDefault();
                      setColResizeState({ col: index, startX: event.clientX, startWidth: colWidths[index] ?? DEFAULT_COL_WIDTH });
                    }}
                    className="absolute right-0 top-0 h-full w-2 cursor-col-resize bg-transparent hover:bg-[#6d7dff]/40"
                  />
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, rowIndex) => (
              <tr key={rowIndex} className="align-top" style={{ height: `${rowHeights[rowIndex] ?? DEFAULT_ROW_HEIGHT}px` }}>
                <td className="relative border-b border-r border-white/10 px-3 py-2 text-xs text-white/50" style={{ height: `${rowHeights[rowIndex] ?? DEFAULT_ROW_HEIGHT}px` }}>
                  {rowIndex + 1}
                  <button
                    type="button"
                    aria-label={`Resize row ${rowIndex + 1}`}
                    onPointerDown={(event) => {
                      event.preventDefault();
                      setRowResizeState({ row: rowIndex, startY: event.clientY, startHeight: rowHeights[rowIndex] ?? DEFAULT_ROW_HEIGHT });
                    }}
                    className="absolute bottom-0 left-0 h-2 w-full cursor-row-resize bg-transparent hover:bg-[#6d7dff]/40"
                  />
                </td>
                {Array.from({ length: maxColumns }).map((_, colIndex) => {
                  const value = row[colIndex] ?? '';
                  const active = selectedCell?.row === rowIndex && selectedCell?.col === colIndex;
                  const inRange = activeRange ? isInRange({ row: rowIndex, col: colIndex }, activeRange) : false;
                  const style = cellStyles[cellKey(rowIndex, colIndex)];
                  return (
                    <td key={`${rowIndex}-${colIndex}`} className="border-b border-r border-white/10 p-0 last:border-r-0" style={{ width: `${colWidths[colIndex] ?? DEFAULT_COL_WIDTH}px`, minWidth: `${colWidths[colIndex] ?? DEFAULT_COL_WIDTH}px`, height: `${rowHeights[rowIndex] ?? DEFAULT_ROW_HEIGHT}px` }}>
                      <textarea
                        value={value}
                        disabled={locked}
                        onMouseDown={() => {
                          const cell = { row: rowIndex, col: colIndex };
                          setDragging(true);
                          setAnchor(cell);
                          setEdge(cell);
                          onSelectCell(cell);
                          
                          // If in formula mode, insert cell reference instead
                          if (activeRange && formulaInput.trim().startsWith('=')) {
                            const cellRef = `${columnLabel(colIndex)}${rowIndex + 1}`;
                            setFormulaInput(formulaInput + cellRef);
                          }
                        }}
                        onMouseEnter={() => {
                          if (!dragging) return;
                          setEdge({ row: rowIndex, col: colIndex });
                        }}
                        onFocus={() => {
                          const cell = { row: rowIndex, col: colIndex };
                          setAnchor(cell);
                          setEdge(cell);
                          onSelectCell(cell);
                        }}
                        onBlur={(event) => {
                          const entered = event.target.value.trim();
                          if (!entered.startsWith('=')) return;
                          onApplyFormulaToRange(entered, { start: { row: rowIndex, col: colIndex }, end: { row: rowIndex, col: colIndex } });
                        }}
                        onMouseUp={(event) => {
                          const nextHeight = event.currentTarget.offsetHeight;
                          const currentRowHeight = rowHeights[rowIndex] ?? DEFAULT_ROW_HEIGHT;
                          if (Math.abs(nextHeight - currentRowHeight) > 1) {
                            onResizeRow(rowIndex, nextHeight);
                          }
                        }}
                        onDoubleClick={() => onResizeRow(rowIndex, estimateAutoRowHeight(value, colWidths[colIndex] ?? DEFAULT_COL_WIDTH))}
                        onChange={(event) => onCellChange(rowIndex, colIndex, event.target.value)}
                        className={`w-full resize-y bg-transparent px-3 py-2 text-sm text-white outline-none ${style?.bold ? 'font-bold' : ''} ${style?.italic ? 'italic' : ''} ${style?.highlight ? 'bg-[#f6c76a]/30' : ''} ${active ? 'ring-2 ring-inset ring-[#6d7dff]/60' : ''} ${inRange ? 'bg-[#6d7dff]/10' : ''}`}
                        style={{ height: `${rowHeights[rowIndex] ?? DEFAULT_ROW_HEIGHT}px`, backgroundColor: style?.highlight ? 'rgba(246, 199, 106, 0.18)' : undefined }}
                      />
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

'use client';

import type { ChangeEvent, PointerEvent as ReactPointerEvent } from 'react';
import { useEffect, useMemo, useRef, useState } from 'react';
import { FileText, Presentation, Table2, Upload, Sparkles, Search, Maximize2, Minimize2, MoonStar, Sun, Sheet, ChevronRight, ChevronDown, Download } from 'lucide-react';
import Link from 'next/link';
import * as mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import { asBlob as htmlAsDocxBlob } from 'html-docx-js-typescript';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';

import type { DocKind, DocumentState, SlideState } from '@/types/documents';
import { detectKind, fileToDataUrl, fileToText, clamp, fallbackSlide } from '@/utils/document-utils';
import { parseAdvancedPptxSlides } from '@/utils/pptx-advanced-parser';
import { parseAdvancedDocx } from '@/utils/docx-advanced-parser';
import { DocxPane } from '@/components/editors/docx-pane';
import { PdfPane } from '@/components/editors/pdf-pane';
import { PptPane } from '@/components/editors/ppt-pane';
import { XlsxPane } from '@/components/editors/xlsx-pane';
import type { CellRange } from '@/components/editors/xlsx-pane';
import { evaluateFormula } from '@/utils/xlsx-formula';

function sheetCellKey(row: number, col: number) {
  return `${row}:${col}`;
}

function parseSheetCellKey(key: string) {
  const [rowText, colText] = key.split(':');
  return { row: Number(rowText), col: Number(colText) };
}

function reindexSheetMap<T>(
  source: Record<string, T>,
  axis: 'row' | 'col',
  removedIndex: number
) {
  const next: Record<string, T> = {};
  for (const [key, value] of Object.entries(source)) {
    const { row, col } = parseSheetCellKey(key);
    if (!Number.isFinite(row) || !Number.isFinite(col)) continue;
    if (axis === 'row') {
      if (row === removedIndex) continue;
      const nextRow = row > removedIndex ? row - 1 : row;
      next[sheetCellKey(nextRow, col)] = value;
    } else {
      if (col === removedIndex) continue;
      const nextCol = col > removedIndex ? col - 1 : col;
      next[sheetCellKey(row, nextCol)] = value;
    }
  }
  return next;
}

function reindexNumericMap(source: Record<number, number>, removedIndex: number) {
  const next: Record<number, number> = {};
  for (const [key, value] of Object.entries(source)) {
    const index = Number(key);
    if (!Number.isFinite(index) || index === removedIndex) continue;
    next[index > removedIndex ? index - 1 : index] = value;
  }
  return next;
}

function normalizeSheetRows(rows: string[][], minRows: number, minCols: number) {
  const targetCols = Math.max(1, minCols, ...rows.map((row) => row.length));
  const normalized = rows.map((row) =>
    Array.from({ length: targetCols }).map((_, colIndex) => row[colIndex] ?? '')
  );
  while (normalized.length < minRows) {
    normalized.push(Array.from({ length: targetCols }).map(() => ''));
  }
  return normalized;
}

function recalculateRowsWithFormulas(baseRows: string[][], formulas: Record<string, string>) {
  let rows = baseRows.map((row) => [...row]);
  const entries = Object.entries(formulas).filter(([, formula]) => formula.trim().startsWith('='));

  for (let pass = 0; pass < 4; pass += 1) {
    let changed = false;
    for (const [key, formula] of entries) {
      const { row, col } = parseSheetCellKey(key);
      if (!Number.isFinite(row) || !Number.isFinite(col) || row < 0 || col < 0) continue;
      rows = normalizeSheetRows(rows, row + 1, col + 1);
      const nextValue = evaluateFormula(formula, rows);
      if (rows[row][col] !== nextValue) {
        rows[row][col] = nextValue;
        changed = true;
      }
    }
    if (!changed) break;
  }

  return rows;
}

type XlsxCellStyle = {
  bold?: boolean;
  italic?: boolean;
  highlight?: boolean;
};

function updateSheetStyles(
  styles: Record<string, XlsxCellStyle>,
  range: CellRange,
  key: keyof XlsxCellStyle
) {
  const top = Math.min(range.start.row, range.end.row);
  const bottom = Math.max(range.start.row, range.end.row);
  const left = Math.min(range.start.col, range.end.col);
  const right = Math.max(range.start.col, range.end.col);
  let allEnabled = true;

  for (let row = top; row <= bottom; row += 1) {
    for (let col = left; col <= right; col += 1) {
      if (!styles[sheetCellKey(row, col)]?.[key]) {
        allEnabled = false;
      }
    }
  }

  const nextStyles = { ...styles };
  for (let row = top; row <= bottom; row += 1) {
    for (let col = left; col <= right; col += 1) {
      const cellKeyValue = sheetCellKey(row, col);
      nextStyles[cellKeyValue] = {
        ...(nextStyles[cellKeyValue] ?? {}),
        [key]: !allEnabled
      };
    }
  }

  return nextStyles;
}

function inferXlsxCellValue(value: string) {
  const trimmed = value.trim();
  if (trimmed === '') return { t: 's' as const, v: '' };
  const numeric = Number(trimmed);
  if (Number.isFinite(numeric)) {
    return { t: 'n' as const, v: numeric };
  }
  if (trimmed.toLowerCase() === 'true' || trimmed.toLowerCase() === 'false') {
    return { t: 'b' as const, v: trimmed.toLowerCase() === 'true' };
  }
  return { t: 's' as const, v: value };
}

function normalizeSlideZIndexes(slide: SlideState) {
  const sorted = [...slide.elements].sort((a, b) => (a.z ?? 0) - (b.z ?? 0));
  return {
    ...slide,
    elements: sorted.map((element, index) => ({ ...element, z: index + 1 }))
  };
}

function buildDocxPdfHtml(contentHtml: string) {
  return `
    <style>
      .docx-export, .docx-export * { box-sizing: border-box; }
      .docx-export {
        font-family: Cambria, Georgia, serif;
        font-size: 12pt;
        line-height: 1.45;
        color: #111827;
        background: #ffffff;
      }
      .docx-export h1,
      .docx-export h2,
      .docx-export h3 {
        line-height: 1.25;
        color: #111827;
        font-weight: 700;
      }
      .docx-export h1 {
        margin: 0 0 1rem;
        font-size: 2rem;
      }
      .docx-export h2 {
        margin: 1.5rem 0 0.75rem;
        font-size: 1.5rem;
      }
      .docx-export h3 {
        margin: 1.25rem 0 0.5rem;
        font-size: 1.2rem;
      }
      .docx-export p {
        margin: 0.75rem 0;
        line-height: 1.75;
      }
      .docx-export ul,
      .docx-export ol {
        margin: 0.75rem 0;
        padding-left: 1.25rem;
      }
      .docx-export li {
        margin: 0.35rem 0;
      }
      .docx-export hr {
        border: 0;
        border-top: 2px dashed rgba(37, 99, 235, 0.6);
        margin: 1.75rem 0;
        position: relative;
      }
      .docx-export hr::after {
        content: 'Page break';
        position: absolute;
        right: 0;
        top: -0.9rem;
        font-size: 0.65rem;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        color: rgba(37, 99, 235, 0.85);
        background: #ffffff;
        padding: 0 0.35rem;
      }
      .docx-export blockquote {
        margin: 1rem 0;
        padding: 0.25rem 0.75rem;
        border-left: 3px solid #9ca3af;
      }
      .docx-export img {
        max-width: 100%;
        height: auto;
      }
    </style>
    <div class="docx-export">${contentHtml}</div>
  `;
}

function buildDocxExportHtml(contentHtml: string) {
  return `<style>
body {
  font-family: Cambria, Georgia, serif;
  font-size: 12pt;
  line-height: 1.45;
  color: #111827;
}
h1 { font-size: 2em; margin: 1em 0 0.5em 0; }
h2 { font-size: 1.5em; margin: 0.8em 0 0.4em 0; }
h3 { font-size: 1.25em; margin: 0.6em 0 0.3em 0; }
p { margin: 0.5em 0; }
ul, ol { margin: 0.5em 0 0.5em 1.25em; }
li { margin: 0.25em 0; }
blockquote { border-left: 3px solid #9ca3af; padding-left: 0.75em; margin: 0.5em 0; }
strong, b { font-weight: bold; }
em, i { font-style: italic; }
u { text-decoration: underline; }
s, strike { text-decoration: line-through; }
</style>${contentHtml || '<p></p>'}`;
}


async function htmlToPdfBlob(html: string) {
  const PX_PER_PT = 96 / 72;
  const PAGE_MARGIN_PT = 28;
  const container = document.createElement('div');
  container.style.position = 'fixed';
  container.style.left = '-99999px';
  container.style.top = '0';
  container.style.width = `${Math.round((595.28 - PAGE_MARGIN_PT * 2) * PX_PER_PT)}px`;
  container.style.background = '#ffffff';
  container.style.padding = '0';
  container.style.color = '#111827';
  container.style.boxSizing = 'border-box';
  container.style.fontFamily = 'Cambria, Georgia, serif';
  container.style.lineHeight = '1.45';
  container.style.fontSize = '12pt';
  container.innerHTML = html;
  document.body.appendChild(container);

  const fullCanvas = await html2canvas(container, {
    scale: 2,
    useCORS: true,
    backgroundColor: '#ffffff'
  });

  document.body.removeChild(container);

  const pdf = new jsPDF('p', 'pt', 'a4');
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const printableWidth = pageWidth - PAGE_MARGIN_PT * 2;
  const printableHeight = pageHeight - PAGE_MARGIN_PT * 2;
  const pageSliceHeightPx = Math.max(1, Math.floor(printableHeight * PX_PER_PT * 2));

  let offsetY = 0;
  let pageIndex = 0;

  while (offsetY < fullCanvas.height) {
    const sliceHeight = Math.min(pageSliceHeightPx, fullCanvas.height - offsetY);
    const pageCanvas = document.createElement('canvas');
    pageCanvas.width = fullCanvas.width;
    pageCanvas.height = sliceHeight;

    const context = pageCanvas.getContext('2d');
    if (!context) {
      throw new Error('Failed to render PDF page context');
    }

    context.drawImage(
      fullCanvas,
      0,
      offsetY,
      fullCanvas.width,
      sliceHeight,
      0,
      0,
      fullCanvas.width,
      sliceHeight
    );

    const pageImage = pageCanvas.toDataURL('image/png');
    const renderHeight = sliceHeight / (PX_PER_PT * 2);

    if (pageIndex > 0) {
      pdf.addPage();
    }

    pdf.addImage(pageImage, 'PNG', PAGE_MARGIN_PT, PAGE_MARGIN_PT, printableWidth, renderHeight, undefined, 'FAST');
    offsetY += sliceHeight;
    pageIndex += 1;
  }

  return pdf.output('blob');
}

export function EditorSuite() {
  const [documents, setDocuments] = useState<DocumentState[]>([]);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [query, setQuery] = useState('');
  const [fileTypeFilter, setFileTypeFilter] = useState<'all' | DocKind>('all');
  const [importTypeFilter, setImportTypeFilter] = useState<'all' | DocKind>('all');
  const [showClearConfirmDialog, setShowClearConfirmDialog] = useState(false);
  const [theme, setTheme] = useState<'dark' | 'light'>('dark');
  const [focusMode, setFocusMode] = useState(false);
  const [statusMessage, setStatusMessage] = useState<string | null>(null);
  const [downloadFormat, setDownloadFormat] = useState<'docx' | 'pdf' | 'csv' | 'xlsx'>('docx');
  const [selectedSlideIndex, setSelectedSlideIndex] = useState(0);
  const [selectedElementId, setSelectedElementId] = useState<string | null>(null);
  const [selectedCell, setSelectedCell] = useState<{ row: number; col: number } | null>(null);
  const [xlsxCalculating, setXlsxCalculating] = useState(false);
  const [dragState, setDragState] = useState<{
    slideIndex: number;
    elementId: string;
    pointerOffsetX: number;
    pointerOffsetY: number;
  } | null>(null);
  const slideStageRef = useRef<HTMLDivElement>(null);
  const documentsRef = useRef<DocumentState[]>([]);

  const selectedDocument = useMemo(
    () => documents.find((document) => document.id === selectedId) ?? documents[0] ?? null,
    [documents, selectedId]
  );

  const filteredDocuments = documents.filter((document) => {
    const needle = query.toLowerCase();
    const matchesText = document.name.toLowerCase().includes(needle) || document.summary.toLowerCase().includes(needle);
    const matchesType = fileTypeFilter === 'all' || document.kind === fileTypeFilter;
    return matchesText && matchesType;
  });

  const updateDocument = (id: string, updater: (document: DocumentState) => DocumentState) => {
    setDocuments((current) => current.map((document) => (document.id === id ? updater(document) : document)));
  };

  const setStatus = (message: string) => {
    setStatusMessage(message);
    window.setTimeout(() => {
      setStatusMessage((current) => (current === message ? null : current));
    }, 2200);
  };

  const getImportAccept = (type: 'all' | DocKind) => {
    if (type === 'docx') return '.doc,.docx';
    if (type === 'pdf') return '.pdf';
    if (type === 'ppt') return '.ppt,.pptx';
    if (type === 'xlsx') return '.xls,.xlsx,.csv';
    return '.doc,.docx,.pdf,.ppt,.pptx,.xls,.xlsx,.csv';
  };

  const revokePreviewUrls = (docs: DocumentState[]) => {
    for (const document of docs) {
      if (document.previewUrl?.startsWith('blob:')) {
        URL.revokeObjectURL(document.previewUrl);
      }
    }
  };

  useEffect(() => {
    const savedTheme = window.localStorage.getItem('office-forge-theme');
    if (savedTheme === 'dark' || savedTheme === 'light') {
      setTheme(savedTheme);
    }
  }, []);

  useEffect(() => {
    document.body.classList.toggle('theme-light', theme === 'light');
    window.localStorage.setItem('office-forge-theme', theme);
    return () => {
      document.body.classList.remove('theme-light');
    };
  }, [theme]);

  useEffect(() => {
    documentsRef.current = documents;
  }, [documents]);

  const clearLibraryWithConfirm = () => {
    if (documents.length === 0) return;
    setShowClearConfirmDialog(true);
  };

  const clearLibraryNow = () => {
    revokePreviewUrls(documents);
    setDocuments([]);
    setSelectedId(null);
    setShowClearConfirmDialog(false);
    setStatus('Library cleared.');
  };

  useEffect(() => {
    const saved = window.localStorage.getItem('office-forge-documents');
    if (!saved) return;
    try {
      const parsed = JSON.parse(saved) as DocumentState[];
      setDocuments(parsed);
      setSelectedId(parsed[0]?.id ?? null);
    } catch {
      window.localStorage.removeItem('office-forge-documents');
    }
  }, []);

  useEffect(() => () => revokePreviewUrls(documentsRef.current), []);

  useEffect(() => {
    try {
      const forStorage = documents.map((document) => ({
        ...document,
        previewUrl: undefined
      }));
      window.localStorage.setItem('office-forge-documents', JSON.stringify(forStorage));
    } catch {
      setStatus('Autosave limited by browser storage.');
    }
  }, [documents]);

  useEffect(() => {
    setSelectedSlideIndex(0);
    setSelectedElementId(null);
    setSelectedCell(null);
  }, [selectedDocument]);

  useEffect(() => {
    if (!selectedDocument) return;
    if (selectedDocument.kind === 'pdf') setDownloadFormat('pdf');
    if (selectedDocument.kind === 'docx') setDownloadFormat('docx');
    if (selectedDocument.kind === 'xlsx') setDownloadFormat('xlsx');
  }, [selectedDocument]);

  useEffect(() => {
    if (!dragState || selectedDocument?.kind !== 'ppt') return;

    const handlePointerMove = (event: PointerEvent) => {
      const stage = slideStageRef.current;
      if (!stage) return;

      const rect = stage.getBoundingClientRect();
      const element = selectedDocument.slides[dragState.slideIndex]?.elements.find((entry) => entry.id === dragState.elementId);
      if (!element) return;

      const nextX = ((event.clientX - rect.left - dragState.pointerOffsetX) / rect.width) * 100;
      const nextY = ((event.clientY - rect.top - dragState.pointerOffsetY) / rect.height) * 100;

      updateDocument(selectedDocument.id, (document) => ({
        ...document,
        slides: document.slides.map((slide, index) => index === dragState.slideIndex ? {
          ...slide,
          elements: slide.elements.map((entry) => entry.id === dragState.elementId ? {
            ...entry,
            x: clamp(nextX, 0, 100 - entry.w),
            y: clamp(nextY, 0, 100 - entry.h)
          } : entry)
        } : slide),
        updatedAt: 'just now'
      }));
    };

    const handlePointerUp = () => setDragState(null);
    window.addEventListener('pointermove', handlePointerMove);
    window.addEventListener('pointerup', handlePointerUp);
    return () => {
      window.removeEventListener('pointermove', handlePointerMove);
      window.removeEventListener('pointerup', handlePointerUp);
    };
  }, [dragState, selectedDocument, selectedDocument?.id, selectedDocument?.kind, selectedDocument?.slides]);

  const handleUpload = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const kind = detectKind(file.name);
    if (importTypeFilter !== 'all' && kind !== importTypeFilter) {
      setStatus(`Selected file is ${kind.toUpperCase()}. Import filter is set to ${importTypeFilter.toUpperCase()}.`);
      event.target.value = '';
      return;
    }

    const id = crypto.randomUUID();
    const created: DocumentState = {
      id,
      name: file.name,
      kind,
      updatedAt: 'just now',
      summary: `Imported ${file.name}.`,
      contentHtml: '',
      slides: [],
      sheets: [],
      pdfNotes: '',
      previewUrl: undefined
    };

    try {
      if (kind === 'docx') {
        // First try advanced parsing to extract images and watermarks
        try {
          const advancedResult = await parseAdvancedDocx(file);
          
          // Still use mammoth for main content conversion (it's better at HTML)
          const mammothResult = await mammoth.convertToHtml({ arrayBuffer: await fileToText(file) });
          created.contentHtml = mammothResult.value || '<p>Document content could not be extracted. The file may be empty or in an unsupported format.</p>';
          
          // Store extracted images and watermark
          if (advancedResult.images.length > 0) {
            created.docxImages = {};
            for (const img of advancedResult.images) {
              created.docxImages[img.id] = img.data;
            }
          }
          if (advancedResult.watermark) {
            created.docxWatermark = advancedResult.watermark;
          }
        } catch (advancedError) {
          // Fall back to basic mammoth parsing if advanced parsing fails
          console.warn('Advanced DOCX parsing failed, using basic conversion:', advancedError);
          const result = await mammoth.convertToHtml({ arrayBuffer: await fileToText(file) });
          created.contentHtml = result.value || '<p>Document content could not be extracted. The file may be empty or in an unsupported format.</p>';
        }
      }

      if (kind === 'xlsx') {
        const workbook = XLSX.read(await fileToText(file), { type: 'array' });
        const sheetName = workbook.SheetNames[0] ?? 'Sheet1';
        const worksheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json<string[]>(worksheet, { header: 1, blankrows: false }) as string[][];
        created.sheets = [{ name: sheetName, rows: rows.length > 0 ? rows.map((row) => row.map((value) => `${value ?? ''}`)) : [['']], formulas: {}, cellStyles: {}, rowHeights: {}, colWidths: {} }];
      }

      if (kind === 'ppt') {
        const slides = await parseAdvancedPptxSlides(file);
        created.slides = slides;
      }

      if (kind === 'pdf') {
        created.previewUrl = URL.createObjectURL(file);
      }
    } catch {
      if (kind === 'ppt') {
        created.slides = [fallbackSlide(file.name.replace(/\.[^.]+$/, ''))];
      }
      if (kind === 'xlsx') {
        created.sheets = [{ name: 'Sheet1', rows: [['']], formulas: {}, cellStyles: {}, rowHeights: {}, colWidths: {} }];
      }
    }

    setDocuments((current) => [created, ...current]);
    setSelectedId(id);
    event.target.value = '';
  };

  const updateCurrentSlide = (updater: (slide: SlideState) => SlideState) => {
    if (!selectedDocument || selectedDocument.kind !== 'ppt') return;
    updateDocument(selectedDocument.id, (document) => ({
      ...document,
      slides: document.slides.map((slide, index) => index === selectedSlideIndex ? updater(slide) : slide),
      updatedAt: 'just now'
    }));
  };

  const addPptTextElement = () => {
    if (!selectedDocument || selectedDocument.kind !== 'ppt') return;
    updateCurrentSlide((slide) => {
      const nextZ = Math.max(0, ...slide.elements.map((element) => element.z ?? 0)) + 1;
      return {
        ...slide,
        elements: [...slide.elements, {
          id: crypto.randomUUID(),
          kind: 'body',
          text: 'New text box',
          x: 20,
          y: 20,
          w: 30,
          h: 18,
          z: nextZ,
          fillColor: '#13203a',
          textColor: '#ffffff'
        }]
      };
    });
  };

  const duplicatePptElement = () => {
    if (!selectedDocument || selectedDocument.kind !== 'ppt' || !selectedElementId) return;
    updateCurrentSlide((slide) => {
      const source = slide.elements.find((element) => element.id === selectedElementId);
      if (!source) return slide;
      const nextZ = Math.max(0, ...slide.elements.map((element) => element.z ?? 0)) + 1;
      return {
        ...slide,
        elements: [...slide.elements, {
          ...source,
          id: crypto.randomUUID(),
          x: clamp(source.x + 2, 0, 100 - source.w),
          y: clamp(source.y + 2, 0, 100 - source.h),
          z: nextZ
        }]
      };
    });
  };

  const deletePptElement = () => {
    if (!selectedDocument || selectedDocument.kind !== 'ppt' || !selectedElementId) return;
    updateCurrentSlide((slide) => {
      const remaining = slide.elements.filter((element) => element.id !== selectedElementId);
      return normalizeSlideZIndexes({ ...slide, elements: remaining });
    });
    setSelectedElementId(null);
  };

  const shiftPptElementLayer = (direction: 1 | -1) => {
    if (!selectedDocument || selectedDocument.kind !== 'ppt' || !selectedElementId) return;
    updateCurrentSlide((slide) => {
      const ordered = [...slide.elements].sort((a, b) => (a.z ?? 0) - (b.z ?? 0));
      const index = ordered.findIndex((element) => element.id === selectedElementId);
      if (index < 0) return slide;
      const swapIndex = index + direction;
      if (swapIndex < 0 || swapIndex >= ordered.length) return slide;
      [ordered[index], ordered[swapIndex]] = [ordered[swapIndex], ordered[index]];
      return normalizeSlideZIndexes({ ...slide, elements: ordered });
    });
  };

  const addPptImageElement = async (file: File) => {
    if (!selectedDocument || selectedDocument.kind !== 'ppt') return;
    const imageSrc = await fileToDataUrl(file);
    updateCurrentSlide((slide) => {
      const nextZ = Math.max(0, ...slide.elements.map((element) => element.z ?? 0)) + 1;
      return {
        ...slide,
        elements: [...slide.elements, {
          id: crypto.randomUUID(),
          kind: 'image',
          text: '',
          x: 18,
          y: 18,
          w: 34,
          h: 28,
          z: nextZ,
          imageSrc
        }]
      };
    });
  };

  const updateSheetCell = (row: number, col: number, value: string) => {
    if (!selectedDocument || selectedDocument.kind !== 'xlsx' || xlsxCalculating) return;
    updateDocument(selectedDocument.id, (document) => ({
      ...document,
      sheets: document.sheets.map((sheet, index) => index === 0 ? {
        ...sheet,
        ...(() => {
          const currentFormulas = { ...(sheet.formulas ?? {}) };
          delete currentFormulas[sheetCellKey(row, col)];
          const normalizedRows = normalizeSheetRows(sheet.rows, row + 1, col + 1);
          normalizedRows[row][col] = value;
          const recalculated = recalculateRowsWithFormulas(normalizedRows, currentFormulas);
          return { rows: recalculated, formulas: currentFormulas };
        })()
      } : sheet),
      updatedAt: 'just now'
    }));
  };

  const applyFormulaToRange = async (value: string, range: CellRange) => {
    if (!selectedDocument || selectedDocument.kind !== 'xlsx') return;
    const top = Math.min(range.start.row, range.end.row);
    const bottom = Math.max(range.start.row, range.end.row);
    const left = Math.min(range.start.col, range.end.col);
    const right = Math.max(range.start.col, range.end.col);

    setXlsxCalculating(true);
    try {
      await new Promise((resolve) => window.setTimeout(resolve, 0));
      updateDocument(selectedDocument.id, (document) => {
        const firstSheet = document.sheets[0] ?? { name: 'Sheet1', rows: [['']], formulas: {}, cellStyles: {}, rowHeights: {}, colWidths: {} };
        const nextFormulas = { ...(firstSheet.formulas ?? {}) };
        const normalizedRows = normalizeSheetRows(firstSheet.rows, bottom + 1, right + 1);

        for (let row = top; row <= bottom; row += 1) {
          for (let col = left; col <= right; col += 1) {
            const key = sheetCellKey(row, col);
            if (value.trim().startsWith('=')) {
              nextFormulas[key] = value;
            } else {
              delete nextFormulas[key];
              normalizedRows[row][col] = value;
            }
          }
        }

        const recalculatedRows = recalculateRowsWithFormulas(normalizedRows, nextFormulas);

        return {
          ...document,
          sheets: document.sheets.map((sheet, index) => index === 0 ? { ...sheet, rows: recalculatedRows, formulas: nextFormulas } : sheet),
          updatedAt: 'just now'
        };
      });
    } finally {
      setXlsxCalculating(false);
    }
  };

  const toggleBoldForRange = (range: CellRange) => {
    if (!selectedDocument || selectedDocument.kind !== 'xlsx') return;
    updateDocument(selectedDocument.id, (document) => {
      const firstSheet = document.sheets[0] ?? { name: 'Sheet1', rows: [['']], formulas: {}, cellStyles: {}, rowHeights: {}, colWidths: {} };
      const nextStyles = updateSheetStyles(firstSheet.cellStyles ?? {}, range, 'bold');

      return {
        ...document,
        sheets: document.sheets.map((sheet, index) => index === 0 ? { ...sheet, cellStyles: nextStyles } : sheet),
        updatedAt: 'just now'
      };
    });
  };

  const toggleItalicForRange = (range: CellRange) => {
    if (!selectedDocument || selectedDocument.kind !== 'xlsx') return;
    updateDocument(selectedDocument.id, (document) => {
      const firstSheet = document.sheets[0] ?? { name: 'Sheet1', rows: [['']], formulas: {}, cellStyles: {}, rowHeights: {}, colWidths: {} };
      const nextStyles = updateSheetStyles(firstSheet.cellStyles ?? {}, range, 'italic');

      return {
        ...document,
        sheets: document.sheets.map((sheet, index) => index === 0 ? { ...sheet, cellStyles: nextStyles } : sheet),
        updatedAt: 'just now'
      };
    });
  };

  const toggleHighlightForRange = (range: CellRange) => {
    if (!selectedDocument || selectedDocument.kind !== 'xlsx') return;
    updateDocument(selectedDocument.id, (document) => {
      const firstSheet = document.sheets[0] ?? { name: 'Sheet1', rows: [['']], formulas: {}, cellStyles: {}, rowHeights: {}, colWidths: {} };
      const nextStyles = updateSheetStyles(firstSheet.cellStyles ?? {}, range, 'highlight');

      return {
        ...document,
        sheets: document.sheets.map((sheet, index) => index === 0 ? { ...sheet, cellStyles: nextStyles } : sheet),
        updatedAt: 'just now'
      };
    });
  };

  const exportCurrentDocument = (targetFormat: 'docx' | 'pdf' | 'csv' | 'xlsx') => {
    if (!selectedDocument) return;

    let content: string = '';
    let extension: string = targetFormat;
    let mime = targetFormat === 'docx'
      ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      : 'application/pdf';

    if (selectedDocument.kind === 'xlsx') {
      extension = targetFormat === 'csv' ? 'csv' : 'xlsx';
      mime = targetFormat === 'csv'
        ? 'text/csv;charset=utf-8'
        : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    } else if (selectedDocument.kind === 'ppt') {
      content = JSON.stringify(selectedDocument.slides, null, 2);
      extension = 'json';
      mime = 'application/json;charset=utf-8';
    }

    const performExport = async () => {
      let blob: Blob;

      if (selectedDocument.kind === 'docx' || selectedDocument.kind === 'pdf') {
        if (targetFormat === 'docx') {
          const docxHtml = buildDocxExportHtml(selectedDocument.contentHtml || '');
          const docxResult = await htmlAsDocxBlob(docxHtml);
          if (docxResult instanceof Blob) {
            blob = docxResult;
          } else if (ArrayBuffer.isView(docxResult)) {
            const typed = new Uint8Array(docxResult.buffer, docxResult.byteOffset, docxResult.byteLength);
            blob = new Blob([new Uint8Array(typed).buffer], { type: mime });
          } else {
            blob = new Blob([String(docxResult)], { type: mime });
          }
        } else {
          const pdfHtml = buildDocxPdfHtml(selectedDocument.contentHtml || '<p></p>');
          blob = await htmlToPdfBlob(pdfHtml);
        }
      } else if (selectedDocument.kind === 'xlsx') {
        const firstSheet = selectedDocument.sheets[0] ?? { name: 'Sheet1', rows: [['']], formulas: {}, cellStyles: {}, rowHeights: {}, colWidths: {} };

        if (targetFormat === 'csv') {
          const csvSheet = XLSX.utils.aoa_to_sheet(firstSheet.rows);
          const csvText = XLSX.utils.sheet_to_csv(csvSheet);
          blob = new Blob([csvText], { type: mime });
        } else {
          const normalizedRows = normalizeSheetRows(firstSheet.rows, firstSheet.rows.length || 1, firstSheet.rows[0]?.length || 1);
          const sheet = XLSX.utils.aoa_to_sheet(normalizedRows);

          for (const [key, formula] of Object.entries(firstSheet.formulas ?? {})) {
            if (!formula.trim().startsWith('=')) continue;
            const { row, col } = parseSheetCellKey(key);
            const address = XLSX.utils.encode_cell({ r: row, c: col });
            const value = normalizedRows[row]?.[col] ?? '';
            const cellValue = inferXlsxCellValue(value);
            sheet[address] = { f: formula.trim().slice(1), v: cellValue.v, t: cellValue.t };
          }

          const rowEntries = normalizedRows.map((_, rowIndex) => {
            const height = firstSheet.rowHeights?.[rowIndex];
            return height ? { hpx: height } : undefined;
          }).filter((entry): entry is { hpx: number } => Boolean(entry));
          const columnCount = Math.max(...normalizedRows.map((row) => row.length), 1);
          const columnEntries = Array.from({ length: columnCount }, (_, columnIndex) => {
            const width = firstSheet.colWidths?.[columnIndex];
            return width ? { wch: Math.max(8, Math.round(width / 8)) } : undefined;
          }).filter((entry): entry is { wch: number } => Boolean(entry));

          if (rowEntries.length > 0) sheet['!rows'] = rowEntries;
          if (columnEntries.length > 0) sheet['!cols'] = columnEntries;

          const workbook = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(workbook, sheet, firstSheet.name || 'Sheet1');
          const workbookData = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
          blob = new Blob([workbookData], { type: mime });
        }
      } else {
        blob = new Blob([content], { type: mime });
      }

      const url = URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      const baseName = selectedDocument.name.replace(/\.[^.]+$/, '') || 'document';
      anchor.href = url;
      anchor.download = `${baseName}-export.${extension}`;
      anchor.click();
      URL.revokeObjectURL(url);
      setStatus('Download ready.');
    };

    performExport().catch(() => setStatus('Export failed.'));
  };

  const sheetRows = selectedDocument?.kind === 'xlsx' ? (selectedDocument.sheets[0]?.rows ?? [['']]) : [['']];
  const sheetFormulas = selectedDocument?.kind === 'xlsx' ? (selectedDocument.sheets[0]?.formulas ?? {}) : {};
  const sheetStyles = selectedDocument?.kind === 'xlsx' ? (selectedDocument.sheets[0]?.cellStyles ?? {}) : {};
  const sheetRowHeights = selectedDocument?.kind === 'xlsx' ? (selectedDocument.sheets[0]?.rowHeights ?? {}) : {};
  const sheetColWidths = selectedDocument?.kind === 'xlsx' ? (selectedDocument.sheets[0]?.colWidths ?? {}) : {};

  const resizeSheetRow = (row: number, height: number) => {
    if (!selectedDocument || selectedDocument.kind !== 'xlsx') return;
    updateDocument(selectedDocument.id, (document) => ({
      ...document,
      sheets: document.sheets.map((sheet, index) => index === 0 ? {
        ...sheet,
        rowHeights: {
          ...(sheet.rowHeights ?? {}),
          [row]: Math.max(36, Math.round(height))
        }
      } : sheet),
      updatedAt: 'just now'
    }));
  };

  const resizeSheetColumn = (col: number, width: number) => {
    if (!selectedDocument || selectedDocument.kind !== 'xlsx') return;
    updateDocument(selectedDocument.id, (document) => ({
      ...document,
      sheets: document.sheets.map((sheet, index) => index === 0 ? {
        ...sheet,
        colWidths: {
          ...(sheet.colWidths ?? {}),
          [col]: Math.max(90, Math.round(width))
        }
      } : sheet),
      updatedAt: 'just now'
    }));
  };

  const addSheetRow = () => {
    if (!selectedDocument || selectedDocument.kind !== 'xlsx') return;
    updateDocument(selectedDocument.id, (document) => ({
      ...document,
      sheets: document.sheets.map((sheet, index) => {
        if (index !== 0) return sheet;
        const targetCols = Math.max(1, ...sheet.rows.map((row) => row.length));
        return {
          ...sheet,
          rows: [...sheet.rows.map((row) => Array.from({ length: targetCols }).map((_, colIndex) => row[colIndex] ?? '')), Array.from({ length: targetCols }).map(() => '')]
        };
      }),
      updatedAt: 'just now'
    }));
  };

  const addSheetColumn = () => {
    if (!selectedDocument || selectedDocument.kind !== 'xlsx') return;
    updateDocument(selectedDocument.id, (document) => ({
      ...document,
      sheets: document.sheets.map((sheet, index) => {
        if (index !== 0) return sheet;
        const rows = sheet.rows.length > 0 ? sheet.rows : [['']];
        const expanded = rows.map((row) => [...row, '']);
        return {
          ...sheet,
          rows: expanded
        };
      }),
      updatedAt: 'just now'
    }));
  };

  const deleteSheetRow = (rowIndex: number) => {
    if (!selectedDocument || selectedDocument.kind !== 'xlsx') return;
    updateDocument(selectedDocument.id, (document) => ({
      ...document,
      sheets: document.sheets.map((sheet, index) => {
        if (index !== 0) return sheet;
        const rows = sheet.rows.length > 0 ? sheet.rows : [['']];
        if (rows.length <= 1) {
          return {
            ...sheet,
            rows: [['']],
            formulas: {},
            cellStyles: {},
            rowHeights: {},
            colWidths: sheet.colWidths ?? {}
          };
        }
        const safeIndex = clamp(rowIndex, 0, rows.length - 1);
        const nextRows = rows.filter((_, indexValue) => indexValue !== safeIndex);
        return {
          ...sheet,
          rows: nextRows,
          formulas: reindexSheetMap(sheet.formulas ?? {}, 'row', safeIndex),
          cellStyles: reindexSheetMap(sheet.cellStyles ?? {}, 'row', safeIndex),
          rowHeights: reindexNumericMap(sheet.rowHeights ?? {}, safeIndex)
        };
      }),
      updatedAt: 'just now'
    }));
    setSelectedCell((current) => {
      if (!current) return null;
      if (current.row === rowIndex) return null;
      if (current.row > rowIndex) return { row: current.row - 1, col: current.col };
      return current;
    });
  };

  const deleteSheetColumn = (colIndex: number) => {
    if (!selectedDocument || selectedDocument.kind !== 'xlsx') return;
    updateDocument(selectedDocument.id, (document) => ({
      ...document,
      sheets: document.sheets.map((sheet, index) => {
        if (index !== 0) return sheet;
        const rows = sheet.rows.length > 0 ? sheet.rows : [['']];
        const maxColumns = Math.max(1, ...rows.map((row) => row.length));
        if (maxColumns <= 1) {
          return {
            ...sheet,
            rows: rows.map(() => ['']),
            formulas: {},
            cellStyles: {},
            colWidths: {},
            rowHeights: sheet.rowHeights ?? {}
          };
        }
        const safeIndex = clamp(colIndex, 0, maxColumns - 1);
        const nextRows = rows.map((row) => {
          const normalized = Array.from({ length: maxColumns }).map((_, indexValue) => row[indexValue] ?? '');
          return normalized.filter((_, indexValue) => indexValue !== safeIndex);
        });
        return {
          ...sheet,
          rows: nextRows,
          formulas: reindexSheetMap(sheet.formulas ?? {}, 'col', safeIndex),
          cellStyles: reindexSheetMap(sheet.cellStyles ?? {}, 'col', safeIndex),
          colWidths: reindexNumericMap(sheet.colWidths ?? {}, safeIndex)
        };
      }),
      updatedAt: 'just now'
    }));
    setSelectedCell((current) => {
      if (!current) return null;
      if (current.col === colIndex) return null;
      if (current.col > colIndex) return { row: current.row, col: current.col - 1 };
      return current;
    });
  };

  return (
    <main className={`min-h-screen p-3 sm:p-4 lg:p-5 ${theme === 'light' ? 'text-[#0f172a]' : 'text-white'}`}>
      <div className="mx-auto flex h-[calc(100vh-1.5rem)] max-w-[1800px] flex-col gap-4 overflow-hidden sm:h-[calc(100vh-2rem)]">
        <section className={`sticky top-0 z-50 shrink-0 overflow-hidden rounded-[28px] border shadow-soft backdrop-blur-xl ${theme === 'light' ? 'border-black/10 bg-white/70' : 'border-white/10 bg-white/5'} ${focusMode ? 'hidden' : ''}`}>
          <div className="flex flex-wrap items-center justify-between gap-2 border-b border-white/10 p-4">
            <div className="flex flex-wrap items-center gap-2">
              <button
                className={`inline-flex items-center gap-2 rounded-full border px-4 py-2.5 text-sm transition ${theme === 'light' ? 'border-black/10 bg-black/[0.03] text-[#0f172a] hover:bg-black/[0.06]' : 'border-white/10 bg-white/5 text-white hover:bg-white/10'}`}
                onClick={clearLibraryWithConfirm}
              >
                <Sparkles className="h-4 w-4" />
                Clear library
              </button>
              <button
                type="button"
                onClick={() => setTheme((current) => current === 'dark' ? 'light' : 'dark')}
                aria-label={theme === 'dark' ? 'Switch to light mode' : 'Switch to dark mode'}
                title={theme === 'dark' ? 'Switch to light mode' : 'Switch to dark mode'}
                className={`inline-flex items-center rounded-full border p-2.5 transition ${theme === 'light' ? 'border-black/10 bg-black/[0.03] text-[#0f172a] hover:bg-black/[0.06]' : 'border-white/10 bg-white/5 text-white hover:bg-white/10'}`}
              >
                {theme === 'dark' ? <Sun className="h-4 w-4" /> : <MoonStar className="h-4 w-4" />}
              </button>
            </div>
            <div className="inline-flex items-center">
              <label className="inline-flex cursor-pointer items-center gap-2 rounded-l-full bg-[#f6c76a] px-4 py-2.5 text-sm font-semibold text-[#111827] transition hover:bg-[#ffe08f]">
                <Upload className="h-4 w-4" />
                Import file
                <input className="hidden" type="file" accept={getImportAccept(importTypeFilter)} onChange={handleUpload} />
              </label>
              <div className="relative">
                <select
                  value={importTypeFilter}
                  onChange={(event) => setImportTypeFilter(event.target.value as 'all' | DocKind)}
                  className="appearance-none rounded-r-full border border-l-0 border-white/10 bg-white/10 py-2.5 pl-3 pr-9 text-sm text-white outline-none"
                  title="Narrow import to a specific type"
                >
                  <option value="all" className="bg-[#07111f]">All</option>
                  <option value="docx" className="bg-[#07111f]">DOCX</option>
                  <option value="pdf" className="bg-[#07111f]">PDF</option>
                  <option value="ppt" className="bg-[#07111f]">PPT</option>
                  <option value="xlsx" className="bg-[#07111f]">XLSX</option>
                </select>
                <ChevronDown className="pointer-events-none absolute right-3 top-1/2 h-4 w-4 -translate-y-1/2 text-white/75" />
              </div>
            </div>
          </div>
        </section>

        <section className={`grid min-h-0 flex-1 gap-4 ${focusMode ? 'lg:grid-cols-[minmax(0,1fr)]' : 'lg:grid-cols-[280px_minmax(0,1fr)]'}`}>
          <aside className={`sticky top-[5.5rem] z-40 min-h-0 overflow-auto rounded-[28px] border border-white/10 bg-[#08101e]/90 backdrop-blur-xl ${focusMode ? 'hidden' : ''}`}>
            <div className="border-b border-white/10 p-4">
              <div className="flex items-center justify-between gap-2">
                <div className="text-sm font-semibold text-white">Library</div>
                <select
                  value={fileTypeFilter}
                  onChange={(event) => setFileTypeFilter(event.target.value as 'all' | DocKind)}
                  className="rounded-full border border-white/10 bg-white/5 px-3 py-1.5 text-xs uppercase tracking-[0.18em] text-white outline-none"
                  title="Filter file type"
                >
                  <option value="all" className="bg-[#07111f]">All</option>
                  <option value="docx" className="bg-[#07111f]">DOCX</option>
                  <option value="pdf" className="bg-[#07111f]">PDF</option>
                  <option value="ppt" className="bg-[#07111f]">PPT</option>
                  <option value="xlsx" className="bg-[#07111f]">XLSX</option>
                </select>
              </div>
              <div className="mt-4 flex items-center gap-2 rounded-2xl border border-white/10 bg-white/5 px-3 py-2 text-white/60">
                <Search className="h-4 w-4" />
                <input
                  value={query}
                  onChange={(event) => setQuery(event.target.value)}
                  placeholder="Find specific file"
                  className="w-full bg-transparent text-sm outline-none placeholder:text-white/30"
                />
              </div>
            </div>
            <div className="space-y-2 overflow-auto p-3">
              {filteredDocuments.map((document) => {
                const Icon = document.kind === 'docx' ? FileText : document.kind === 'ppt' ? Presentation : document.kind === 'xlsx' ? Table2 : Sheet;
                const active = document.id === selectedDocument?.id;
                return (
                  <button
                    key={document.id}
                    onClick={() => setSelectedId(document.id)}
                    className={`flex w-full items-center gap-3 rounded-[22px] border p-3 text-left transition ${active ? 'border-[#6d7dff]/50 bg-[#6d7dff]/15' : 'border-transparent bg-white/[0.04] hover:bg-white/[0.07]'}`}
                  >
                    <div className="rounded-2xl bg-white/10 p-3 text-white">
                      <Icon className="h-5 w-5" />
                    </div>
                    <div className="min-w-0 flex-1">
                      <div className="flex items-center justify-between gap-3">
                        <div className="truncate text-sm font-medium text-white">{document.name}</div>
                        <ChevronRight className="h-4 w-4 text-white/35" />
                      </div>
                      <div className="mt-2 text-[11px] uppercase tracking-[0.24em] text-white/35">{document.updatedAt}</div>
                    </div>
                  </button>
                );
              })}
            </div>
          </aside>

          <section className="grid min-h-0 overflow-hidden rounded-[28px] border border-white/10 bg-[#08101e]/92 shadow-soft backdrop-blur-xl">
            <div className="min-h-0 overflow-hidden">
              <div className="flex flex-wrap items-center justify-between gap-3 border-b border-white/10 p-4 sm:p-5">
                <div>
                  {selectedDocument ? (
                    <div className="mt-1 flex flex-wrap items-center gap-2 text-lg font-semibold text-white">
                      {selectedDocument.name}
                      <span className="rounded-full border border-white/10 bg-white/5 px-2 py-0.5 text-[11px] uppercase tracking-[0.24em] text-white/45">{selectedDocument.kind}</span>
                    </div>
                  ) : (
                    <div className="mt-1 text-sm text-white/55">No file loaded.</div>
                  )}
                </div>
                <div className="flex flex-wrap gap-2">
                  {selectedDocument?.kind === 'docx' && (
                    <select
                      value={downloadFormat}
                      onChange={(event) => setDownloadFormat(event.target.value as 'docx' | 'pdf' | 'csv' | 'xlsx')}
                      className="rounded-full border border-white/10 bg-white/5 px-3 py-2 text-sm text-white outline-none"
                    >
                      <option value="docx" className="bg-[#07111f]">DOCX</option>
                      <option value="pdf" className="bg-[#07111f]">PDF</option>
                    </select>
                  )}
                  {selectedDocument?.kind === 'xlsx' && (
                    <select
                      value={downloadFormat}
                      onChange={(event) => setDownloadFormat(event.target.value as 'docx' | 'pdf' | 'csv' | 'xlsx')}
                      className="rounded-full border border-white/10 bg-white/5 px-3 py-2 text-sm text-white outline-none"
                    >
                      <option value="xlsx" className="bg-[#07111f]">XLSX</option>
                      <option value="csv" className="bg-[#07111f]">CSV</option>
                    </select>
                  )}
                  {selectedDocument && (
                    <button
                      className="rounded-full border border-white/10 bg-white/5 px-4 py-2 text-sm text-white/90 transition hover:bg-white/10"
                      onClick={() => exportCurrentDocument(
                        selectedDocument.kind === 'pdf'
                          ? 'pdf'
                          : selectedDocument.kind === 'docx'
                            ? downloadFormat
                            : selectedDocument.kind === 'xlsx'
                              ? downloadFormat
                              : 'docx'
                      )}
                    >
                      <Download className="mr-2 inline h-4 w-4" />
                      Download
                    </button>
                  )}
                  <button className="rounded-full border border-white/10 bg-white/5 px-4 py-2 text-sm text-white/80 transition hover:bg-white/10" onClick={() => {
                    setFocusMode((current) => !current);
                    setStatus(!focusMode ? 'Focus mode enabled.' : 'Focus mode disabled.');
                  }}>
                    {focusMode ? <Minimize2 className="mr-2 inline h-4 w-4" /> : <Maximize2 className="mr-2 inline h-4 w-4" />}
                    {focusMode ? 'Exit focus' : 'Focus mode'}
                  </button>
                </div>
              </div>

              <div className="h-[calc(100%-5.6rem)] min-h-0 overflow-auto p-4 sm:p-5">
                {!selectedDocument && (
                  <div className="rounded-[24px] border border-dashed border-white/20 bg-[#07111f] p-10 text-center">
                    <div className="text-lg font-semibold text-white">No file loaded</div>
                  </div>
                )}

                {selectedDocument?.kind === 'docx' && (
                  <DocxPane
                    html={selectedDocument.contentHtml}
                    watermark={selectedDocument.docxWatermark}
                    images={selectedDocument.docxImages}
                    onChange={(nextHtml) => updateDocument(selectedDocument.id, (document) => ({
                      ...document,
                      contentHtml: nextHtml,
                      updatedAt: 'just now'
                    }))}
                  />
                )}

                {selectedDocument?.kind === 'pdf' && (
                  <PdfPane
                    name={selectedDocument.name}
                    previewUrl={selectedDocument.previewUrl}
                  />
                )}

                {selectedDocument?.kind === 'ppt' && (
                  <PptPane
                    name={selectedDocument.name}
                    slides={selectedDocument.slides}
                    selectedSlideIndex={selectedSlideIndex}
                    selectedElementId={selectedElementId}
                    slideStageRef={slideStageRef}
                    onSelectSlide={(index) => {
                      setSelectedSlideIndex(index);
                      setSelectedElementId(selectedDocument.slides[index]?.elements[0]?.id ?? null);
                    }}
                    onAddSlide={() => {
                      updateDocument(selectedDocument.id, (document) => ({
                        ...document,
                        slides: [...document.slides, fallbackSlide(`Slide ${document.slides.length + 1}`)],
                        updatedAt: 'just now'
                      }));
                      setSelectedSlideIndex(selectedDocument.slides.length);
                    }}
                    onAddTextBox={addPptTextElement}
                    onDuplicateElement={duplicatePptElement}
                    onDeleteElement={deletePptElement}
                    onBringForward={() => shiftPptElementLayer(1)}
                    onSendBackward={() => shiftPptElementLayer(-1)}
                    onAddImage={(file) => {
                      addPptImageElement(file).catch(() => setStatus('Image import failed.'));
                    }}
                    onElementPointerDown={(event: ReactPointerEvent<HTMLDivElement>, elementId: string) => {
                      if (!slideStageRef.current) return;
                      const slide = selectedDocument.slides[selectedSlideIndex];
                      const element = slide?.elements.find((entry) => entry.id === elementId);
                      if (!element) return;
                      const stageRect = slideStageRef.current.getBoundingClientRect();
                      const offsetX = event.clientX - stageRect.left - (element.x / 100) * stageRect.width;
                      const offsetY = event.clientY - stageRect.top - (element.y / 100) * stageRect.height;
                      setSelectedElementId(element.id);
                      setDragState({ slideIndex: selectedSlideIndex, elementId: element.id, pointerOffsetX: offsetX, pointerOffsetY: offsetY });
                      event.currentTarget.setPointerCapture(event.pointerId);
                    }}
                    onElementTextChange={(elementId, value) => updateCurrentSlide((slide) => ({
                      ...slide,
                      elements: slide.elements.map((element) => element.id === elementId ? { ...element, text: value } : element)
                    }))}
                    onSlideUpdate={(index, updates) => {
                      updateDocument(selectedDocument.id, (document) => ({
                        ...document,
                        slides: document.slides.map((slide, slideIndex) => 
                          slideIndex === index ? { ...slide, ...updates } : slide
                        ),
                        updatedAt: 'just now'
                      }));
                    }}
                  />
                )}

                {selectedDocument?.kind === 'xlsx' && (
                  <XlsxPane
                    rows={sheetRows}
                    formulas={sheetFormulas}
                    cellStyles={sheetStyles}
                    rowHeights={sheetRowHeights}
                    colWidths={sheetColWidths}
                    selectedCell={selectedCell}
                    locked={xlsxCalculating}
                    onSelectCell={setSelectedCell}
                    onCellChange={updateSheetCell}
                    onApplyFormulaToRange={applyFormulaToRange}
                    onToggleBold={toggleBoldForRange}
                    onToggleItalic={toggleItalicForRange}
                    onToggleHighlight={toggleHighlightForRange}
                    onAddRow={addSheetRow}
                    onAddColumn={addSheetColumn}
                    onDeleteRow={deleteSheetRow}
                    onDeleteColumn={deleteSheetColumn}
                    onResizeRow={resizeSheetRow}
                    onResizeColumn={resizeSheetColumn}
                  />
                )}
              </div>
            </div>
          </section>
        </section>

        {statusMessage && (
          <div className="pointer-events-none fixed bottom-4 right-4 z-20 rounded-full border border-white/15 bg-[#07111f]/95 px-4 py-2 text-sm text-white shadow-soft">
            {statusMessage}
          </div>
        )}

        {showClearConfirmDialog && (
          <div className="fixed inset-0 z-40 flex items-center justify-center bg-[#030712]/70 p-4 backdrop-blur-sm">
            <div className="w-full max-w-md rounded-2xl border border-white/15 bg-[#07111f] p-5 shadow-soft">
              <div className="text-lg font-semibold text-white">Clear library?</div>
              <p className="mt-2 text-sm text-white/70">This removes all imported files from the current session and cannot be undone.</p>
              <div className="mt-5 flex justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setShowClearConfirmDialog(false)}
                  className="rounded-full border border-white/15 bg-white/5 px-4 py-2 text-sm text-white/85 transition hover:bg-white/10"
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={clearLibraryNow}
                  className="rounded-full border border-[#f87171]/60 bg-[#7f1d1d]/40 px-4 py-2 text-sm font-semibold text-[#fecaca] transition hover:bg-[#991b1b]/50"
                >
                  Clear library
                </button>
              </div>
            </div>
          </div>
        )}
      </div>

      <footer className={`mx-auto mt-12 flex max-w-[1800px] items-center justify-between gap-4 rounded-2xl border px-4 py-4 text-sm sm:mt-16 ${theme === 'light' ? 'border-black/10 bg-white/65 text-[#0f172a]' : 'border-white/10 bg-white/5 text-white/80'}`}>
        <span>Copyright © 2026 Office Forge. Built for open access document editing.</span>
        <Link
          href="/about"
          className={`rounded-full border px-4 py-1.5 transition ${theme === 'light' ? 'border-black/10 bg-black/[0.03] text-[#0f172a] hover:bg-black/[0.06]' : 'border-white/10 bg-white/5 text-white hover:bg-white/10'}`}
        >
          About
        </Link>
      </footer>
    </main>
  );
}

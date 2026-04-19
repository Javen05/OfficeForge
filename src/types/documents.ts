export type DocKind = 'docx' | 'pdf' | 'ppt' | 'xlsx';

export type SheetState = {
  name: string;
  rows: string[][];
  formulas?: Record<string, string>;
  cellStyles?: Record<string, { bold?: boolean; italic?: boolean; highlight?: boolean }>;
  rowHeights?: Record<number, number>;
  colWidths?: Record<number, number>;
};

export type SlideElementState = {
  id: string;
  kind: 'title' | 'body' | 'note' | 'image';
  text: string;
  x: number;
  y: number;
  w: number;
  h: number;
  z?: number;
  fillColor?: string;
  textColor?: string;
  imageSrc?: string;
};

export type SlideState = {
  title: string;
  bullets: string[];
  elements: SlideElementState[];
  backgroundColor?: string;
  backgroundImage?: string;
  speakerNotes?: string;
  transitionType?: string;
  layoutName?: string;
};

export type DocumentState = {
  id: string;
  name: string;
  kind: DocKind;
  updatedAt: string;
  summary: string;
  contentHtml: string;
  slides: SlideState[];
  sheets: SheetState[];
  pdfNotes: string;
  previewUrl?: string;
  // DOCX specific fields
  docxImages?: Record<string, string>; // id -> base64 data URL
  docxWatermark?: string;
};

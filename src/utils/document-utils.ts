import type { DocKind, SlideElementState, SlideState } from '@/types/documents';

export function fileToText(file: File) {
  return file.arrayBuffer();
}

export function fileToDataUrl(file: File) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(typeof reader.result === 'string' ? reader.result : '');
    reader.onerror = () => reject(new Error('Failed to read file data URL'));
    reader.readAsDataURL(file);
  });
}

export function detectKind(name: string): DocKind {
  const lower = name.toLowerCase();
  if (lower.endsWith('.pdf')) return 'pdf';
  if (lower.endsWith('.ppt') || lower.endsWith('.pptx')) return 'ppt';
  if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) return 'xlsx';
  return 'docx';
}

export function createSlideElements(title: string): SlideElementState[] {
  return [
    { id: crypto.randomUUID(), kind: 'title', text: title, x: 8, y: 12, w: 56, h: 16, z: 1, fillColor: '#ffffff', textColor: '#13203a' },
    { id: crypto.randomUUID(), kind: 'body', text: 'Edit content and placement directly on the slide.', x: 8, y: 34, w: 54, h: 30, z: 2, fillColor: '#13203a', textColor: '#ffffff' },
    { id: crypto.randomUUID(), kind: 'note', text: 'Drag each card to reposition it.', x: 66, y: 64, w: 22, h: 16, z: 3, fillColor: '#f6c76a', textColor: '#13203a' }
  ];
}

export function fallbackSlide(title: string): SlideState {
  return {
    title,
    bullets: [],
    elements: createSlideElements(title)
  };
}

export function clamp(value: number, min: number, max: number) {
  return Math.min(max, Math.max(min, value));
}

export function columnLabel(index: number) {
  let value = index;
  let label = '';
  while (value >= 0) {
    label = String.fromCharCode((value % 26) + 65) + label;
    value = Math.floor(value / 26) - 1;
  }
  return label;
}

import JSZip from 'jszip';
import type { SlideState } from '@/types/documents';
import { createSlideElements, fallbackSlide } from '@/utils/document-utils';

function decodeXmlEntities(value: string) {
  return value
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

function toPercent(value: number, total: number) {
  if (!total || !Number.isFinite(total) || total <= 0) return 0;
  return (value / total) * 100;
}

function readNumber(value?: string) {
  const parsed = Number(value ?? '0');
  return Number.isFinite(parsed) ? parsed : 0;
}

function getMimeFromPath(path: string) {
  const lower = path.toLowerCase();
  if (lower.endsWith('.png')) return 'image/png';
  if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg';
  if (lower.endsWith('.gif')) return 'image/gif';
  if (lower.endsWith('.webp')) return 'image/webp';
  if (lower.endsWith('.bmp')) return 'image/bmp';
  if (lower.endsWith('.svg')) return 'image/svg+xml';
  return 'application/octet-stream';
}

function joinTargetPath(basePath: string, target: string) {
  if (target.startsWith('/')) {
    return target.replace(/^\//, '');
  }

  const baseSegments = basePath.split('/').slice(0, -1);
  const targetSegments = target.split('/');
  const resolved: string[] = [...baseSegments];

  for (const segment of targetSegments) {
    if (!segment || segment === '.') continue;
    if (segment === '..') {
      resolved.pop();
      continue;
    }
    resolved.push(segment);
  }

  return resolved.join('/');
}

function extractShapeText(shapeXml: string) {
  const withBreaks = shapeXml.replace(/<a:br\s*\/>/g, '\n');
  const paragraphs = Array.from(shapeXml.matchAll(/<a:p[\s\S]*?<\/a:p>/g));
  if (paragraphs.length === 0) {
    const singleRun = Array.from(withBreaks.matchAll(/<(?:a:t|a:fld[^>]*)>([\s\S]*?)<\/(?:a:t|a:fld)>/g))
      .map((entry) => decodeXmlEntities(entry[1]))
      .join('')
      .trim();
    return singleRun;
  }

  const lines = paragraphs
    .map((paragraph) =>
      Array.from(paragraph[0].matchAll(/<a:t>([\s\S]*?)<\/a:t>/g))
        .map((entry) => decodeXmlEntities(entry[1]))
        .join('')
        .trim()
    )
    .filter(Boolean);

  return lines.join('\n');
}

function parseColor(xml: string | undefined) {
  if (!xml) return undefined;
  const srgb = xml.match(/<a:srgbClr[^>]*val="([A-Fa-f0-9]{6})"/);
  if (srgb?.[1]) return `#${srgb[1]}`;
  return undefined;
}

function parseShapeStyle(shapeXml: string) {
  const fillColor = parseColor(shapeXml.match(/<a:solidFill[\s\S]*?<\/a:solidFill>/)?.[0]);
  const textColor = parseColor(shapeXml.match(/<a:rPr[\s\S]*?<\/a:rPr>/)?.[0]);
  return { fillColor, textColor };
}

function parseOffExt(xml: string, slideWidth: number, slideHeight: number) {
  const offMatch = xml.match(/<a:off[^>]*x="([\d-]+)"[^>]*y="([\d-]+)"/);
  const extMatch = xml.match(/<a:ext[^>]*cx="([\d-]+)"[^>]*cy="([\d-]+)"/);
  return {
    x: toPercent(readNumber(offMatch?.[1]), slideWidth),
    y: toPercent(readNumber(offMatch?.[2]), slideHeight),
    w: Math.max(4, toPercent(readNumber(extMatch?.[1]), slideWidth)),
    h: Math.max(4, toPercent(readNumber(extMatch?.[2]), slideHeight))
  };
}

function extractSlideBackgroundColor(slideXml: string) {
  const srgb = slideXml.match(/<p:bgPr[\s\S]*?<a:solidFill>[\s\S]*?<a:srgbClr[^>]*val="([A-Fa-f0-9]{6})"/);
  if (srgb?.[1]) {
    return `#${srgb[1]}`;
  }
  return undefined;
}

export async function parsePptxSlides(file: File): Promise<SlideState[]> {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());

  let slideWidth = 12192000;
  let slideHeight = 6858000;
  const presentationXml = await zip.file('ppt/presentation.xml')?.async('text');
  if (presentationXml) {
    const sizeMatch = presentationXml.match(/<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"/);
    if (sizeMatch) {
      slideWidth = readNumber(sizeMatch[1]) || slideWidth;
      slideHeight = readNumber(sizeMatch[2]) || slideHeight;
    }
  }

  const slidePaths = Object.keys(zip.files)
    .filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
    .sort((a, b) => {
      const aNum = Number(a.match(/slide(\d+)\.xml$/)?.[1] ?? 0);
      const bNum = Number(b.match(/slide(\d+)\.xml$/)?.[1] ?? 0);
      return aNum - bNum;
    });

  if (slidePaths.length === 0) {
    return [fallbackSlide(file.name.replace(/\.[^.]+$/, ''))];
  }

  const slides: SlideState[] = [];

  for (const path of slidePaths) {
    const xml = await zip.file(path)?.async('text');
    if (!xml) continue;

    const relPath = `ppt/slides/_rels/${path.split('/').pop()}.rels`;
    const relXml = await zip.file(relPath)?.async('text');
    const relMap = new Map<string, string>();

    if (relXml) {
      for (const relMatch of relXml.matchAll(/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*>/g)) {
        if (/TargetMode="External"/.test(relMatch[0])) continue;
        relMap.set(relMatch[1], joinTargetPath(relPath, relMatch[2]));
      }
    }

    const elements: SlideState['elements'] = [];
    let z = 1;

    const shapeBlocks = Array.from(xml.matchAll(/<p:sp[\s\S]*?<\/p:sp>/g));
    for (const block of shapeBlocks) {
      const shapeXml = block[0];
      const text = extractShapeText(shapeXml);
      if (!text) continue;

      const pos = parseOffExt(shapeXml, slideWidth, slideHeight);
      const style = parseShapeStyle(shapeXml);

      const placeholderType = shapeXml.match(/<p:ph[^>]*type="([^"]+)"/)?.[1] ?? '';
      const kind = placeholderType.includes('title') ? 'title' : 'body';

      elements.push({
        id: crypto.randomUUID(),
        kind,
        text,
        x: pos.x,
        y: pos.y,
        w: pos.w,
        h: pos.h,
        z: z += 1,
        fillColor: style.fillColor,
        textColor: style.textColor
      });
    }

    const frameBlocks = Array.from(xml.matchAll(/<p:graphicFrame[\s\S]*?<\/p:graphicFrame>/g));
    for (const block of frameBlocks) {
      const frameXml = block[0];
      const text = extractShapeText(frameXml);
      if (!text) continue;
      const pos = parseOffExt(frameXml, slideWidth, slideHeight);
      elements.push({
        id: crypto.randomUUID(),
        kind: 'body',
        text,
        x: pos.x,
        y: pos.y,
        w: pos.w,
        h: pos.h,
        z: z += 1
      });
    }

    const pictureBlocks = Array.from(xml.matchAll(/<p:pic[\s\S]*?<\/p:pic>/g));
    for (const block of pictureBlocks) {
      const picXml = block[0];
      const pos = parseOffExt(picXml, slideWidth, slideHeight);
      const embedId = picXml.match(/<a:blip[^>]*r:embed="([^"]+)"/)?.[1];
      const mediaPath = embedId ? relMap.get(embedId) : undefined;
      if (!mediaPath) continue;

      const mediaFile = zip.file(mediaPath);
      if (!mediaFile) continue;
      const base64 = await mediaFile.async('base64');
      const imageSrc = `data:${getMimeFromPath(mediaPath)};base64,${base64}`;

      elements.push({
        id: crypto.randomUUID(),
        kind: 'image',
        text: '',
        x: pos.x,
        y: pos.y,
        w: pos.w,
        h: pos.h,
        z: z += 1,
        imageSrc
      });
    }

    const allTexts = elements.filter((entry) => entry.kind !== 'image').map((entry) => entry.text).filter(Boolean);
    const titleElement = elements.find((entry) => entry.kind === 'title' && entry.text.trim().length > 0);
    const title = titleElement?.text.split('\n')[0] ?? allTexts[0] ?? path.split('/').pop()?.replace('.xml', '') ?? 'Slide';
    const bullets = allTexts.slice(1, 8);

    slides.push({
      title,
      bullets,
      elements: elements.length > 0 ? elements : createSlideElements(title),
      backgroundColor: extractSlideBackgroundColor(xml)
    });
  }

  return slides.length > 0 ? slides : [fallbackSlide(file.name.replace(/\.[^.]+$/, ''))];
}

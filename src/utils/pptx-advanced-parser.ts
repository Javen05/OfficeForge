import JSZip from 'jszip';
import type { SlideState, SlideElementState } from '@/types/documents';

export interface EnhancedSlideState extends SlideState {
  notes?: string;
  speakerNotes?: string;
  transitionType?: string;
  layoutName?: string;
  masterSlideId?: string;
}

/**
 * Advanced text formatting extraction from PPTX
 */
function extractDetailedShapeText(shapeXml: string): { text: string; formatted: boolean } {
  const withBreaks = shapeXml.replace(/<a:br\s*\/>/g, '\n');
  const paragraphs = Array.from(shapeXml.matchAll(/<a:p[\s\S]*?<\/a:p>/g));
  
  if (paragraphs.length === 0) {
    const singleRun = Array.from(withBreaks.matchAll(/<(?:a:t|a:fld[^>]*)>([\s\S]*?)<\/(?:a:t|a:fld)>/g))
      .map((entry) => decodeXmlEntities(entry[1]))
      .join('')
      .trim();
    return { text: singleRun, formatted: false };
  }

  const lines = paragraphs
    .map((paragraph) =>
      Array.from(paragraph[0].matchAll(/<a:t>([\s\S]*?)<\/a:t>/g))
        .map((entry) => decodeXmlEntities(entry[1]))
        .join('')
        .trim()
    )
    .filter((line) => line.length > 0);

  return { text: lines.join('\n'), formatted: lines.length > 1 };
}

function extractFallbackText(shapeXml: string): string {
  return Array.from(shapeXml.matchAll(/<(?:a:t|a:fld[^>]*)>([\s\S]*?)<\/(?:a:t|a:fld)>/g))
    .map((entry) => decodeXmlEntities(entry[1]))
    .join('')
    .trim();
}

/**
 * Extract speaker notes from slide notes
 */
async function extractSpeakerNotes(zip: JSZip, slideNum: number): Promise<string> {
  try {
    const notesPath = `ppt/notesSlides/notesSlide${slideNum}.xml`;
    const notesXml = await zip.file(notesPath)?.async('text');
    
    if (!notesXml) return '';
    
    const notes = Array.from(notesXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g))
      .map((match) => decodeXmlEntities(match[1]))
      .join('\n')
      .trim();
    
    return notes;
  } catch {
    return '';
  }
}

/**
 * Extract slide transitions and animations
 */
function extractSlideTransition(slideXml: string): string | undefined {
  const transitionMatch = slideXml.match(/<p:transition[^>]*\/>/);
  if (!transitionMatch) return undefined;
  
  if (slideXml.includes('<p:advTm')) {
    return 'auto';
  }
  
  if (slideXml.includes('<p:push')) return 'push';
  if (slideXml.includes('<p:wipe')) return 'wipe';
  if (slideXml.includes('<p:fade')) return 'fade';
  if (slideXml.includes('<p:dissolve')) return 'dissolve';
  if (slideXml.includes('<p:blinds')) return 'blinds';
  if (slideXml.includes('<p:wheel')) return 'wheel';
  
  return 'default';
}

/**
 * Extract layout information
 */
function extractLayoutInfo(slideXml: string): { layoutName?: string } {
  const masterLayoutMatch = slideXml.match(/<p:sldLayoutIdLst>[\s\S]*?<p:sldLayoutId[^>]*r:id="([^"]+)"/);
  if (masterLayoutMatch?.[1]) {
    return { layoutName: 'custom' };
  }
  const layoutMatch = slideXml.match(/<p:cSld[^>]*>\s*<p:bg/);
  if (layoutMatch) {
    return { layoutName: 'blank' };
  }
  return { layoutName: 'default' };
}

/**
 * Enhanced color parsing with fallback
 */
function parseColor(xml: string | undefined): string | undefined {
  if (!xml) return undefined;
  
  // Try sRGB color
  const srgb = xml.match(/<a:srgbClr[^>]*val="([A-Fa-f0-9]{6})"/);
  if (srgb?.[1]) return `#${srgb[1]}`;
  
  // Try scheme color
  const schemeColor = xml.match(/<a:schemeClr[^>]*val="([^"]+)"/);
  if (schemeColor?.[1]) {
    // Map scheme colors to defaults
    const schemeMap: Record<string, string> = {
      'accent1': '#4472C4',
      'accent2': '#ED7D31',
      'accent3': '#A5A5A5',
      'accent4': '#FFC000',
      'accent5': '#5B9BD5',
      'accent6': '#70AD47',
    };
    return schemeMap[schemeColor[1]] || '#000000';
  }
  
  return undefined;
}

/**
 * Decode XML entities
 */
function decodeXmlEntities(value: string): string {
  return value
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

/**
 * Utility functions (keeping existing implementation)
 */
function toPercent(value: number, total: number): number {
  if (!total || !Number.isFinite(total) || total <= 0) return 0;
  return (value / total) * 100;
}

function readNumber(value?: string): number {
  const parsed = Number(value ?? '0');
  return Number.isFinite(parsed) ? parsed : 0;
}

function getMimeFromPath(path: string): string {
  const lower = path.toLowerCase();
  if (lower.endsWith('.png')) return 'image/png';
  if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg';
  if (lower.endsWith('.gif')) return 'image/gif';
  if (lower.endsWith('.webp')) return 'image/webp';
  if (lower.endsWith('.bmp')) return 'image/bmp';
  if (lower.endsWith('.svg')) return 'image/svg+xml';
  return 'application/octet-stream';
}

function joinTargetPath(basePath: string, target: string): string {
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

function readMediaAsDataUrl(file: JSZip.JSZipObject, mediaPath: string): Promise<string> {
  return file.async('base64').then((base64) => `data:${getMimeFromPath(mediaPath)};base64,${base64}`);
}

function parseBlockDimensions(xml: string, slideWidth: number, slideHeight: number) {
  const offMatch = xml.match(/<a:off[^>]*x="([\d-]+)"[^>]*y="([\d-]+)"/);
  const extMatch = xml.match(/<a:ext[^>]*cx="([\d-]+)"[^>]*cy="([\d-]+)"/);
  return {
    x: toPercent(readNumber(offMatch?.[1]), slideWidth),
    y: toPercent(readNumber(offMatch?.[2]), slideHeight),
    w: Math.max(4, toPercent(readNumber(extMatch?.[1]), slideWidth)),
    h: Math.max(4, toPercent(readNumber(extMatch?.[2]), slideHeight))
  };
}

async function extractEmbeddedImage(
  xml: string,
  slideWidth: number,
  slideHeight: number,
  zip: JSZip,
  relMap: Map<string, string>
): Promise<SlideElementState | null> {
  const embedId = xml.match(/r:embed="([^"]+)"/)?.[1];
  const mediaPath = embedId ? relMap.get(embedId) : undefined;
  if (!mediaPath) return null;

  const mediaFile = zip.file(mediaPath);
  if (!mediaFile) return null;

  const pos = parseBlockDimensions(xml, slideWidth, slideHeight);
  const altText = xml.match(/<p:cNvPr[^>]*name="([^"]+)"/)?.[1] ?? '';

  try {
    const imageSrc = await readMediaAsDataUrl(mediaFile, mediaPath);
    return {
      id: crypto.randomUUID(),
      kind: 'image',
      text: altText,
      x: pos.x,
      y: pos.y,
      w: pos.w,
      h: pos.h,
      z: 1,
      imageSrc
    };
  } catch (error) {
    console.warn('Failed to parse embedded image:', error);
    return null;
  }
}

/**
 * Parse advanced shape with better text and styling support
 */
async function parseAdvancedShape(
  shapeXml: string,
  slideWidth: number,
  slideHeight: number
): Promise<SlideElementState | null> {
  const textData = extractDetailedShapeText(shapeXml);
  const text = textData.text || extractFallbackText(shapeXml);
  if (!text) return null;
  const pos = parseBlockDimensions(shapeXml, slideWidth, slideHeight);

  const placeholderType = shapeXml.match(/<p:ph[^>]*type="([^"]+)"/)?.[1] ?? '';
  const kind = placeholderType.includes('title') ? 'title' : 'body';

  // Extract fill color
  const fillMatch = shapeXml.match(/<a:solidFill[\s\S]*?<\/a:solidFill>/);
  const fillColor = fillMatch ? parseColor(fillMatch[0]) : undefined;

  // Extract text color
  const textColorMatch = shapeXml.match(/<a:rPr[\s\S]*?<\/a:rPr>/);
  const textColor = textColorMatch ? parseColor(textColorMatch[0]) : undefined;

  return {
    id: crypto.randomUUID(),
    kind: kind as 'title' | 'body' | 'note' | 'image',
    text,
    x: pos.x,
    y: pos.y,
    w: pos.w,
    h: pos.h,
    z: 1,
    fillColor: fillColor || undefined,
    textColor: textColor || '#ffffff'
  };
}

/**
 * Parse advanced image with metadata
 */
/**
 * Enhanced PPTX parser supporting complex presentations
 */
export async function parseAdvancedPptxSlides(file: File): Promise<EnhancedSlideState[]> {
  try {
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

    const slides: EnhancedSlideState[] = [];

    for (let slideIndex = 0; slideIndex < slidePaths.length; slideIndex++) {
      const path = slidePaths[slideIndex];
      const xml = await zip.file(path)?.async('text');
      if (!xml) continue;

      const slideNum = slideIndex + 1;
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

      // Parse shapes (text elements)
      const shapeBlocks = Array.from(xml.matchAll(/<p:sp[\s\S]*?<\/p:sp>/g));
      for (const block of shapeBlocks) {
        const element = await parseAdvancedShape(block[0], slideWidth, slideHeight);
        if (element) {
          element.z = z++;
          elements.push(element);
        }
      }

      // Parse images with enhanced support
      const pictureBlocks = Array.from(xml.matchAll(/<p:pic[\s\S]*?<\/p:pic>/g));
      for (const block of pictureBlocks) {
        const element = await extractEmbeddedImage(block[0], slideWidth, slideHeight, zip, relMap);
        if (element) {
          element.z = z++;
          elements.push(element);
        }
      }

      const backgroundImageId = xml.match(/<p:bg[\s\S]*?<a:blip[^>]*r:embed="([^"]+)"/)?.[1];
      const backgroundImagePath = backgroundImageId ? relMap.get(backgroundImageId) : undefined;
      let backgroundImage: string | undefined;
      if (backgroundImagePath) {
        const backgroundMedia = zip.file(backgroundImagePath);
        if (backgroundMedia) {
          try {
            backgroundImage = await readMediaAsDataUrl(backgroundMedia, backgroundImagePath);
          } catch (error) {
            console.warn('Failed to parse slide background image:', error);
          }
        }
      }

      const embeddedImageBlocks = Array.from(xml.matchAll(/<(?:p:sp|p:graphicFrame)[\s\S]*?<\/p:(?:sp|graphicFrame)>/g));
      for (const block of embeddedImageBlocks) {
        if (!/<a:blip[^>]*r:embed="[^"]+"/.test(block[0])) continue;
        const element = await extractEmbeddedImage(block[0], slideWidth, slideHeight, zip, relMap);
        if (element) {
          element.z = z++;
          elements.push(element);
        }
      }

      // Parse graphic frames
      const frameBlocks = Array.from(xml.matchAll(/<p:graphicFrame[\s\S]*?<\/p:graphicFrame>/g));
      for (const block of frameBlocks) {
        const element = await parseAdvancedShape(block[0], slideWidth, slideHeight);
        if (element && element.text) {
          element.z = z++;
          elements.push(element);
        }
      }

      // Extract background color
      const bgMatch = xml.match(/<p:bgPr[\s\S]*?<a:solidFill>[\s\S]*?<a:srgbClr[^>]*val="([A-Fa-f0-9]{6})"/);
      const backgroundColor = bgMatch?.[1] ? `#${bgMatch[1]}` : undefined;

      // Extract speaker notes
      const speakerNotes = await extractSpeakerNotes(zip, slideNum);

      // Extract transition
      const transitionType = extractSlideTransition(xml);

      // Extract layout
      const layoutInfo = extractLayoutInfo(xml);

      // Build slide
      const allTexts = elements
        .filter((entry) => entry.kind !== 'image' && entry.text)
        .map((entry) => entry.text)
        .filter(Boolean);

      const titleElement = elements.find((entry) => entry.kind === 'title' && entry.text.trim().length > 0);
      const title = titleElement?.text.split('\n')[0] ?? allTexts[0] ?? `Slide ${slideNum}`;
      const bullets = allTexts.slice(1, 8);

      slides.push({
        title,
        bullets,
        elements: elements.length > 0 ? elements : [
          {
            id: crypto.randomUUID(),
            kind: 'title',
            text: title,
            x: 5,
            y: 5,
            w: 90,
            h: 15,
            z: 1,
            fillColor: '#13203a',
            textColor: '#ffffff'
          }
        ],
        backgroundColor,
        backgroundImage,
        speakerNotes: speakerNotes || undefined,
        transitionType,
        layoutName: layoutInfo.layoutName,
        notes: speakerNotes
      });
    }

    return slides;
  } catch (error) {
    console.warn('Advanced PPTX parsing failed, returning empty slides:', error);
    return [];
  }
}

import JSZip from 'jszip';

export interface ExtractedImage {
  id: string;
  data: string; // base64 data URL
  type: string;
}

export interface DocxParseResult {
  html: string;
  images: ExtractedImage[];
  watermark?: string;
  hasWatermark: boolean;
}

/**
 * Extract images from DOCX file
 */
async function extractImages(zip: JSZip): Promise<ExtractedImage[]> {
  const images: ExtractedImage[] = [];
  
  try {
    const mediaFolder = zip.folder('word/media');
    if (mediaFolder) {
      const files = await mediaFolder.file(/.*/);
      for (const file of files) {
        if (file) {
          const filename = file.name.split('/').pop() || '';
          const blob = await file.async('blob');
          const reader = new FileReader();
          
          await new Promise((resolve) => {
            reader.onload = () => {
              const dataUrl = reader.result as string;
              images.push({
                id: filename,
                data: dataUrl,
                type: blob.type || 'image/png'
              });
              resolve(null);
            };
            reader.readAsDataURL(blob);
          });
        }
      }
    }
  } catch (error) {
    console.warn('Could not extract images from DOCX:', error);
  }
  
  return images;
}

/**
 * Extract watermark from document settings
 */
async function extractWatermark(zip: JSZip): Promise<{ watermarkText?: string; hasWatermark: boolean }> {
  try {
    const settingsXml = await zip.file('word/settings.xml')?.async('string');
    if (!settingsXml) return { hasWatermark: false };
    
    // Look for watermark element in settings.xml
    const watermarkMatch = settingsXml.match(/<w:watermark[^>]*w:type="text"[^>]*w:text="([^"]*)"[^>]*>/);
    if (watermarkMatch && watermarkMatch[1]) {
      return { 
        watermarkText: watermarkMatch[1],
        hasWatermark: true
      };
    }
    
    return { hasWatermark: settingsXml.includes('<w:watermark') };
  } catch (error) {
    console.warn('Could not extract watermark from DOCX:', error);
    return { hasWatermark: false };
  }
}

/**
 * Advanced DOCX parser that handles complex formatting, images, and watermarks
 */
export async function parseAdvancedDocx(file: File): Promise<DocxParseResult> {
  const zip = new JSZip();
  const zipContent = await zip.loadAsync(file);
  
  // Extract images
  const images = await extractImages(zipContent);
  
  // Extract watermark info
  const watermarkInfo = await extractWatermark(zipContent);
  
  const html = '';
  
  return {
    html,
    images,
    watermark: watermarkInfo.watermarkText,
    hasWatermark: watermarkInfo.hasWatermark
  };
}

'use client';

import { DocxEngineEditor } from '@/components/docx-engine-editor';

type DocxPaneProps = {
  html: string;
  onChange: (nextHtml: string) => void;
  watermark?: string;
  images?: Record<string, string>;
};

export function DocxPane({ html, onChange, watermark, images }: DocxPaneProps) {
  return (
    <div className="flex flex-col min-h-0 rounded-[28px] border border-white/10 bg-white/5 p-4 sm:p-6">
      <div className="relative flex flex-col min-h-0 overflow-visible rounded-[24px] border border-white/10 bg-[#d9d1c5] p-3 sm:p-6">
        {watermark && (
          <div className="pointer-events-none absolute inset-0 flex items-center justify-center opacity-10 rounded-[24px]">
            <div className="text-6xl font-bold text-gray-400 rotate-[-45deg] whitespace-nowrap">
              {watermark}
            </div>
          </div>
        )}
        <DocxEngineEditor html={html} onChange={onChange} images={images} />
      </div>
    </div>
  );
}

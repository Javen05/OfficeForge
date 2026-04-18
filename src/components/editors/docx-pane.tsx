'use client';

import { DocxEngineEditor } from '@/components/docx-engine-editor';

type DocxPaneProps = {
  html: string;
  onChange: (nextHtml: string) => void;
};

export function DocxPane({ html, onChange }: DocxPaneProps) {
  return (
    <div className="rounded-[28px] border border-white/10 bg-white/5 p-4 sm:p-6">
      <div className="overflow-visible rounded-[24px] border border-white/10 bg-[#d9d1c5] p-3 sm:p-6">
        <DocxEngineEditor html={html} onChange={onChange} />
      </div>
    </div>
  );
}

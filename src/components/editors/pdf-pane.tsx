'use client';

type PdfPaneProps = {
  name: string;
  previewUrl?: string;
};

export function PdfPane({ name, previewUrl }: PdfPaneProps) {
  return (
    <div className="rounded-[24px] border border-dashed border-white/15 bg-[#07111f] p-4">
      {previewUrl ? (
        <iframe
          title={name}
          src={`${previewUrl}#toolbar=1&navpanes=1&scrollbar=1&view=FitH`}
          className="h-[calc(100vh-300px)] min-h-[520px] w-full rounded-[18px] border border-white/10 bg-black"
        />
      ) : (
        <div className="flex h-[calc(100vh-300px)] min-h-[520px] items-center justify-center rounded-[18px] border border-white/10 bg-black/30 px-6 text-center text-sm text-white/65">
          Re-upload this PDF to open it in the browser native viewer.
        </div>
      )}
    </div>
  );
}

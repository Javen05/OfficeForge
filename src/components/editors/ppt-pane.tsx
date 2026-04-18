'use client';

import { useEffect, useRef } from 'react';
import type { PointerEvent as ReactPointerEvent } from 'react';
import { GripVertical } from 'lucide-react';
import type { SlideState } from '@/types/documents';

type PptPaneProps = {
  name: string;
  slides: SlideState[];
  selectedSlideIndex: number;
  selectedElementId: string | null;
  slideStageRef: React.RefObject<HTMLDivElement | null>;
  onSelectSlide: (index: number) => void;
  onAddSlide: () => void;
  onAddTextBox: () => void;
  onDuplicateElement: () => void;
  onDeleteElement: () => void;
  onBringForward: () => void;
  onSendBackward: () => void;
  onAddImage: (file: File) => void;
  onElementPointerDown: (event: ReactPointerEvent<HTMLDivElement>, elementId: string) => void;
  onElementTextChange: (elementId: string, value: string) => void;
};

export function PptPane({
  name,
  slides,
  selectedSlideIndex,
  selectedElementId,
  slideStageRef,
  onSelectSlide,
  onAddSlide,
  onAddTextBox,
  onDuplicateElement,
  onDeleteElement,
  onBringForward,
  onSendBackward,
  onAddImage,
  onElementPointerDown,
  onElementTextChange
}: PptPaneProps) {
  const imageInputRef = useRef<HTMLInputElement>(null);
  useEffect(() => {
    if (slides.length > 0 && (selectedSlideIndex < 0 || selectedSlideIndex >= slides.length)) {
      onSelectSlide(0);
    }
  }, [onSelectSlide, selectedSlideIndex, slides.length]);

  if (slides.length === 0) {
    return (
      <div className="rounded-[24px] border border-white/10 bg-[#07111f] p-6 text-white">
        <div className="text-lg font-semibold">No slides available</div>
        <p className="mt-2 text-sm text-white/70">Add a slide to start editing your presentation.</p>
        <button type="button" onClick={onAddSlide} className="mt-4 rounded-full border border-white/10 bg-white/5 px-4 py-2 text-sm text-white transition hover:bg-white/10">Add slide</button>
      </div>
    );
  }

  const currentSlide = slides[Math.min(selectedSlideIndex, slides.length - 1)];

  const orderedElements = [...currentSlide.elements].sort((a, b) => (a.z ?? 0) - (b.z ?? 0));

  return (
    <div className="grid min-h-0 gap-4 xl:grid-cols-[220px_minmax(0,1fr)]">
      <div className="min-h-0 overflow-auto rounded-[24px] border border-white/10 bg-[#07111f] p-3">
        <div className="space-y-2">
          {slides.map((slide, index) => {
            const active = index === selectedSlideIndex;
            return (
              <button
                key={`${slide.title}-${index}`}
                onClick={() => onSelectSlide(index)}
                className={`w-full rounded-[18px] border p-3 text-left transition ${active ? 'border-[#6d7dff]/50 bg-[#6d7dff]/15' : 'border-white/10 bg-white/5 hover:bg-white/10'}`}
              >
                <div className="text-xs uppercase tracking-[0.24em] text-white/35">Slide {index + 1}</div>
                <div className="mt-1 text-sm font-semibold text-white">{slide.title}</div>
              </button>
            );
          })}
        </div>
        <button className="mt-3 w-full rounded-full border border-white/10 bg-white/5 px-3 py-2 text-sm text-white/85 transition hover:bg-white/10" onClick={onAddSlide}>
          Add slide
        </button>
      </div>

      <div className="min-h-0 space-y-4 rounded-[24px] border border-white/10 bg-[#07111f] p-4 sm:p-5">
        <div className="flex flex-wrap gap-2">
          <button type="button" onClick={onAddTextBox} className="rounded-full border border-white/10 bg-white/5 px-3 py-1.5 text-xs text-white transition hover:bg-white/10">Add text</button>
          <button type="button" onClick={() => imageInputRef.current?.click()} className="rounded-full border border-white/10 bg-white/5 px-3 py-1.5 text-xs text-white transition hover:bg-white/10">Add image</button>
          <button type="button" onClick={onDuplicateElement} className="rounded-full border border-white/10 bg-white/5 px-3 py-1.5 text-xs text-white transition hover:bg-white/10">Duplicate</button>
          <button type="button" onClick={onDeleteElement} className="rounded-full border border-white/10 bg-white/5 px-3 py-1.5 text-xs text-white transition hover:bg-white/10">Delete</button>
          <button type="button" onClick={onBringForward} className="rounded-full border border-white/10 bg-white/5 px-3 py-1.5 text-xs text-white transition hover:bg-white/10">Bring forward</button>
          <button type="button" onClick={onSendBackward} className="rounded-full border border-white/10 bg-white/5 px-3 py-1.5 text-xs text-white transition hover:bg-white/10">Send backward</button>
          <input
            ref={imageInputRef}
            type="file"
            accept="image/*"
            className="hidden"
            onChange={(event) => {
              const file = event.target.files?.[0];
              if (file) onAddImage(file);
              event.currentTarget.value = '';
            }}
          />
        </div>
        <div className="rounded-[24px] border border-white/10 bg-[#0f1b31] p-3 sm:p-4">
          <div
            ref={slideStageRef}
            className="relative aspect-[16/9] w-full overflow-hidden rounded-[20px] shadow-inner"
            style={{ background: currentSlide.backgroundColor ?? 'linear-gradient(135deg,#f9f7f2_0%,#ece3d4_100%)' }}
          >
            <div className="absolute inset-x-4 top-4 flex items-center justify-between text-[10px] uppercase tracking-[0.3em] text-black/30">
              <span>{name}</span>
              <span>Editable stage</span>
            </div>
            {orderedElements.map((element) => {
              const active = element.id === selectedElementId;
              const kindStyle = element.kind === 'image'
                ? 'bg-transparent text-[#13203a]'
                : element.kind === 'title'
                  ? 'bg-white text-[#13203a]'
                  : element.kind === 'body'
                    ? 'bg-[#13203a] text-white'
                    : 'bg-[#f6c76a] text-[#13203a]';
              return (
                <div
                  key={element.id}
                  onPointerDown={(event) => {
                    const target = event.target as HTMLElement;
                    if (target.closest('textarea')) return;
                    onElementPointerDown(event, element.id);
                  }}
                  className={`absolute rounded-[16px] border p-3 text-left shadow-md transition ${active ? 'border-[#6d7dff] ring-2 ring-[#6d7dff]/30' : 'border-black/10'} ${kindStyle}`}
                  style={{
                    left: `${element.x}%`,
                    top: `${element.y}%`,
                    width: `${element.w}%`,
                    height: `${element.h}%`,
                    zIndex: element.z ?? 1,
                    backgroundColor: element.kind === 'image' ? undefined : element.fillColor
                  }}
                >
                  <GripVertical className="pointer-events-none absolute right-2 top-2 h-3.5 w-3.5 opacity-70" />
                  {element.kind === 'image' && element.imageSrc ? (
                    <img
                      src={element.imageSrc}
                      alt="Slide asset"
                      draggable={false}
                      className="h-full w-full rounded-[10px] object-contain"
                    />
                  ) : (
                    <textarea
                      value={element.text}
                      onPointerDown={(event) => event.stopPropagation()}
                      onChange={(event) => onElementTextChange(element.id, event.target.value)}
                      className="h-full w-full resize-none bg-transparent text-sm leading-6 outline-none"
                      style={{ color: element.textColor }}
                    />
                  )}
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </div>
  );
}

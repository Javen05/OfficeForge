'use client';

import { useEffect, useRef, useState } from 'react';
import type { PointerEvent as ReactPointerEvent } from 'react';
import { GripVertical, FileText, Settings2 } from 'lucide-react';
import Image from 'next/image';
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
  onSlideUpdate?: (index: number, updates: Partial<SlideState>) => void;
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
  onElementTextChange,
  onSlideUpdate
}: PptPaneProps) {
  const imageInputRef = useRef<HTMLInputElement>(null);
  const [showNotes, setShowNotes] = useState(false);
  const [showSlideProperties, setShowSlideProperties] = useState(false);
  
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
    <div className="grid min-h-0 gap-4 xl:grid-cols-[220px_minmax(0,1fr)_240px]">
      {/* Slide thumbnails sidebar */}
      <div className="flex min-h-0 flex-col overflow-hidden rounded-[24px] border border-white/10 bg-[#07111f] p-3">
        <div className="min-h-0 flex-1 space-y-2 overflow-auto pr-1">
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
        <button className="mt-3 w-full shrink-0 rounded-full border border-white/10 bg-white/5 px-3 py-2 text-sm text-white/85 transition hover:bg-white/10" onClick={onAddSlide}>
          Add slide
        </button>
      </div>

      {/* Main editing area */}
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
            style={{
              backgroundColor: currentSlide.backgroundColor ?? '#f3ede2',
              backgroundImage: currentSlide.backgroundImage ? `url(${currentSlide.backgroundImage})` : 'linear-gradient(135deg,#f9f7f2_0%,#ece3d4_100%)',
              backgroundSize: currentSlide.backgroundImage ? 'cover' : 'auto',
              backgroundPosition: 'center',
              backgroundRepeat: 'no-repeat'
            }}
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
                    backgroundColor: element.kind === 'image' ? undefined : element.fillColor,
                    overflow: 'hidden'
                  }}
                >
                  <GripVertical className="pointer-events-none absolute right-2 top-2 h-3.5 w-3.5 opacity-70" />
                  {element.kind === 'image' && element.imageSrc ? (
                    <Image
                      src={element.imageSrc}
                      alt="Slide asset"
                      fill
                      unoptimized
                      draggable={false}
                      className="rounded-[10px] object-contain"
                    />
                  ) : (
                    <div className="flex h-full w-full items-start overflow-auto">
                      <textarea
                        value={element.text}
                        onPointerDown={(event) => event.stopPropagation()}
                        onChange={(event) => onElementTextChange(element.id, event.target.value)}
                        className="min-h-full w-full resize-none bg-transparent text-sm leading-6 outline-none overflow-hidden whitespace-pre-wrap break-words"
                        style={{ color: element.textColor }}
                      />
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>

        <div className="rounded-[20px] border border-white/10 bg-white/5 p-3">
          <button
            type="button"
            onClick={() => setShowSlideProperties((current) => !current)}
            className={`flex w-full items-center gap-2 rounded-[12px] border px-3 py-2 text-sm transition ${
              showSlideProperties ? 'border-[#6d7dff]/50 bg-[#6d7dff]/15 text-white' : 'border-white/10 bg-white/5 text-white/70 hover:bg-white/10'
            }`}
          >
            <Settings2 className="h-4 w-4" />
            Slide Properties
          </button>

          {showSlideProperties && onSlideUpdate && (
            <div className="mt-3 space-y-3 rounded-[16px] border border-white/10 bg-[#07111f] p-3">
              <div>
                <label className="mb-1 block text-[11px] uppercase tracking-[0.2em] text-white/40">Title</label>
                <input
                  value={currentSlide.title}
                  onChange={(event) => onSlideUpdate(selectedSlideIndex, { title: event.target.value })}
                  className="w-full rounded-[12px] border border-white/10 bg-white/5 px-3 py-2 text-sm text-white outline-none focus:border-[#6d7dff]/50 focus:bg-white/10"
                />
              </div>
              <div>
                <label className="mb-1 block text-[11px] uppercase tracking-[0.2em] text-white/40">Background color</label>
                <input
                  value={currentSlide.backgroundColor ?? ''}
                  onChange={(event) => onSlideUpdate(selectedSlideIndex, { backgroundColor: event.target.value || undefined })}
                  placeholder="#f3ede2"
                  className="w-full rounded-[12px] border border-white/10 bg-white/5 px-3 py-2 text-sm text-white outline-none focus:border-[#6d7dff]/50 focus:bg-white/10"
                />
              </div>
              <div>
                <label className="mb-1 block text-[11px] uppercase tracking-[0.2em] text-white/40">Layout</label>
                <input
                  value={currentSlide.layoutName ?? ''}
                  onChange={(event) => onSlideUpdate(selectedSlideIndex, { layoutName: event.target.value || undefined })}
                  placeholder="default"
                  className="w-full rounded-[12px] border border-white/10 bg-white/5 px-3 py-2 text-sm text-white outline-none focus:border-[#6d7dff]/50 focus:bg-white/10"
                />
              </div>
              <div>
                <label className="mb-1 block text-[11px] uppercase tracking-[0.2em] text-white/40">Transition</label>
                <input
                  value={currentSlide.transitionType ?? ''}
                  onChange={(event) => onSlideUpdate(selectedSlideIndex, { transitionType: event.target.value || undefined })}
                  placeholder="fade"
                  className="w-full rounded-[12px] border border-white/10 bg-white/5 px-3 py-2 text-sm text-white outline-none focus:border-[#6d7dff]/50 focus:bg-white/10"
                />
              </div>
              <div className="rounded-[12px] border border-white/10 bg-white/5 p-2 text-xs text-white/60">
                <div className="font-semibold text-white/80">Slide summary</div>
                <div>{orderedElements.length} objects on this slide</div>
                {currentSlide.backgroundImage && <div className="mt-1 break-all text-white/50">Background image loaded</div>}
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Speaker notes and slide properties sidebar */}
      <div className="min-h-0 space-y-4 overflow-auto rounded-[24px] border border-white/10 bg-[#07111f] p-4">
        <div className="space-y-2">
          <button
            onClick={() => setShowNotes(!showNotes)}
            className={`w-full flex items-center gap-2 rounded-[12px] border px-3 py-2 text-sm transition ${
              showNotes ? 'border-[#6d7dff]/50 bg-[#6d7dff]/15 text-white' : 'border-white/10 bg-white/5 text-white/70 hover:bg-white/10'
            }`}
          >
            <FileText className="h-4 w-4" />
            Speaker Notes
          </button>

          {showNotes && (
            <div className="space-y-2">
              <textarea
                value={currentSlide.speakerNotes || ''}
                onChange={(e) => {
                  if (onSlideUpdate) {
                    onSlideUpdate(selectedSlideIndex, { speakerNotes: e.target.value });
                  }
                }}
                placeholder="Add speaker notes for this slide..."
                className="w-full h-32 rounded-[12px] border border-white/10 bg-white/5 p-2 text-xs text-white placeholder-white/40 outline-none focus:border-[#6d7dff]/50 focus:bg-white/10 resize-none"
              />
            </div>
          )}

          <div className="rounded-[12px] border border-white/10 bg-white/5 p-2 text-xs text-white/60">
            <div className="font-semibold text-white/80">Layout</div>
            <div className="capitalize">{currentSlide.layoutName ?? 'default'}</div>
          </div>

          <div className="rounded-[12px] border border-white/10 bg-white/5 p-2 text-xs text-white/60">
            <div className="font-semibold text-white/80">Transition</div>
            <div className="capitalize">{currentSlide.transitionType ?? 'none'}</div>
          </div>

          <div className="rounded-[12px] border border-white/10 bg-white/5 p-2 text-xs text-white/60">
            <div className="font-semibold text-white/80">Elements</div>
            <div>{orderedElements.length} objects on this slide</div>
          </div>
        </div>
      </div>
    </div>
  );
}

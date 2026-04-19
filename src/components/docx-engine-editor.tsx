'use client';

import { useEffect, useState } from 'react';
import { EditorContent, useEditor } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import TextAlign from '@tiptap/extension-text-align';
import Underline from '@tiptap/extension-underline';
import Highlight from '@tiptap/extension-highlight';
import { Color, TextStyle } from '@tiptap/extension-text-style';
import Link from '@tiptap/extension-link';
import Subscript from '@tiptap/extension-subscript';
import Superscript from '@tiptap/extension-superscript';

type DocxEngineEditorProps = {
  html: string;
  onChange: (nextHtml: string) => void;
  images?: Record<string, string>;
};

type ToolbarButtonProps = {
  label: string;
  onClick: () => void;
  active?: boolean;
};

function ToolbarButton({ label, onClick, active }: ToolbarButtonProps) {
  return (
    <button
      type="button"
      onClick={onClick}
      className={`rounded-full border px-3 py-1.5 text-xs transition ${active ? 'border-[#6d7dff]/70 bg-[#6d7dff]/20 text-white' : 'border-white/20 bg-white/5 text-white/80 hover:bg-white/10'}`}
    >
      {label}
    </button>
  );
}

export function DocxEngineEditor({ html, onChange, images }: DocxEngineEditorProps) {
  const [showPageBreaks, setShowPageBreaks] = useState(true);
  const [toolbarExpanded, setToolbarExpanded] = useState(true);
  const [linkEditorOpen, setLinkEditorOpen] = useState(false);
  const [linkDraft, setLinkDraft] = useState('https://');

  const editor = useEditor({
    extensions: [
      StarterKit.configure({
        heading: {
          levels: [1, 2, 3]
        }
      }),
      TextAlign.configure({
        types: ['heading', 'paragraph']
      }),
      Underline,
      Highlight,
      TextStyle,
      Color,
      Link.configure({
        openOnClick: false,
        autolink: true,
        defaultProtocol: 'https'
      }),
      Subscript,
      Superscript
    ],
    content: html || '<h1>Start typing</h1><p>Your DOCX content will keep its layout here.</p>',
    editorProps: {
      attributes: {
        class: 'tiptap-docx prose prose-slate max-w-none min-h-[520px] rounded-[18px] bg-white p-6 text-[#151515] outline-none sm:p-10'
      }
    },
    onUpdate: ({ editor: current }) => {
      onChange(current.getHTML());
    }
  });

  // Global CSS for enhanced document rendering
  useEffect(() => {
    const style = document.createElement('style');
    style.textContent = `
      .tiptap-docx img {
        max-width: 100%;
        height: auto;
        margin: 0.5em 0;
        border-radius: 4px;
      }
      .tiptap-docx table {
        border-collapse: collapse;
        width: 100%;
        margin: 1em 0;
      }
      .tiptap-docx td, .tiptap-docx th {
        border: 1px solid #d1d5db;
        padding: 0.5em;
      }
      .tiptap-docx th {
        background-color: #f3f4f6;
        font-weight: bold;
      }
    `;
    document.head.appendChild(style);
    return () => style.remove();
  }, []);

  useEffect(() => {
    if (!editor) return;
    
    // Embed extracted images into the HTML
    let processedHtml = html;
    if (images) {
      for (const [id, dataUrl] of Object.entries(images)) {
        // Replace image references with data URLs
        processedHtml = processedHtml.replace(
          new RegExp(`src="[^"]*${id}"`, 'g'),
          `src="${dataUrl}"`
        );
        // Also handle cases where images might be referenced differently
        processedHtml = processedHtml.replace(
          new RegExp(`src="[^"]*${id.split('.')[0]}`, 'g'),
          `src="${dataUrl}"`
        );
      }
    }
    
    const current = editor.getHTML();
    if (current !== processedHtml) {
      editor.commands.setContent(processedHtml || '<h1>Start typing</h1><p>Your DOCX content will keep its layout here.</p>', {
        emitUpdate: false
      });
    }
  }, [editor, html, images]);

  if (!editor) {
    return <div className="min-h-[520px] rounded-[18px] bg-white p-6 text-[#151515]">Loading editor...</div>;
  }

  const openLinkEditor = () => {
    const previous = editor.getAttributes('link').href as string | undefined;
    setLinkDraft(previous || 'https://');
    setLinkEditorOpen(true);
  };

  const applyLink = () => {
    if (!linkDraft.trim()) {
      editor.chain().focus().unsetLink().run();
      setLinkEditorOpen(false);
      return;
    }
    editor.chain().focus().setLink({ href: linkDraft.trim() }).run();
    setLinkEditorOpen(false);
  };

  return (
    <div className="flex flex-col min-h-0 gap-3">
      <div className="sticky top-0 z-50 space-y-2 rounded-2xl border border-white/10 bg-[#07111f]/95 p-3 shadow-soft backdrop-blur">
        <div className={`flex flex-wrap items-center gap-2 ${toolbarExpanded ? 'border-b border-white/10 pb-2' : ''}`}>
          <span className="mr-auto text-[10px] uppercase tracking-[0.16em] text-white/45">DOCX toolbar</span>
          <ToolbarButton label={toolbarExpanded ? 'Collapse' : 'Expand'} onClick={() => setToolbarExpanded((current) => !current)} active={toolbarExpanded} />
        </div>

        {toolbarExpanded && (
          <>
        <div className="flex flex-wrap items-center gap-2 border-b border-white/10 pb-2">
          <span className="mr-1 text-[10px] uppercase tracking-[0.16em] text-white/45">Quick</span>
          <ToolbarButton label="Undo" onClick={() => editor.chain().focus().undo().run()} />
          <ToolbarButton label="Redo" onClick={() => editor.chain().focus().redo().run()} />
          <ToolbarButton label="Clear" onClick={() => editor.chain().focus().unsetAllMarks().clearNodes().run()} />
        </div>

        <div className="flex flex-wrap items-center gap-2 border-b border-white/10 pb-2">
          <span className="mr-1 text-[10px] uppercase tracking-[0.16em] text-white/45">Text</span>
          <ToolbarButton label="B" onClick={() => editor.chain().focus().toggleBold().run()} active={editor.isActive('bold')} />
          <ToolbarButton label="I" onClick={() => editor.chain().focus().toggleItalic().run()} active={editor.isActive('italic')} />
          <ToolbarButton label="U" onClick={() => editor.chain().focus().toggleUnderline().run()} active={editor.isActive('underline')} />
          <ToolbarButton label="S" onClick={() => editor.chain().focus().toggleStrike().run()} active={editor.isActive('strike')} />
          <ToolbarButton label="H1" onClick={() => editor.chain().focus().toggleHeading({ level: 1 }).run()} active={editor.isActive('heading', { level: 1 })} />
          <ToolbarButton label="H2" onClick={() => editor.chain().focus().toggleHeading({ level: 2 }).run()} active={editor.isActive('heading', { level: 2 })} />
          <ToolbarButton label="H3" onClick={() => editor.chain().focus().toggleHeading({ level: 3 }).run()} active={editor.isActive('heading', { level: 3 })} />
          <ToolbarButton label="Sup" onClick={() => editor.chain().focus().toggleSuperscript().run()} active={editor.isActive('superscript')} />
          <ToolbarButton label="Sub" onClick={() => editor.chain().focus().toggleSubscript().run()} active={editor.isActive('subscript')} />
          <ToolbarButton label="HL" onClick={() => editor.chain().focus().toggleHighlight().run()} active={editor.isActive('highlight')} />
          <ToolbarButton label="Black" onClick={() => editor.chain().focus().setColor('#111827').run()} />
          <ToolbarButton label="Blue" onClick={() => editor.chain().focus().setColor('#1d4ed8').run()} />
          <ToolbarButton label="Red" onClick={() => editor.chain().focus().setColor('#b91c1c').run()} />
        </div>

        <div className="flex flex-wrap items-center gap-2 border-b border-white/10 pb-2">
          <span className="mr-1 text-[10px] uppercase tracking-[0.16em] text-white/45">Paragraph</span>
          <ToolbarButton label="Bullet" onClick={() => editor.chain().focus().toggleBulletList().run()} active={editor.isActive('bulletList')} />
          <ToolbarButton label="Number" onClick={() => editor.chain().focus().toggleOrderedList().run()} active={editor.isActive('orderedList')} />
          <ToolbarButton label="Quote" onClick={() => editor.chain().focus().toggleBlockquote().run()} active={editor.isActive('blockquote')} />
          <ToolbarButton label="Code" onClick={() => editor.chain().focus().toggleCodeBlock().run()} active={editor.isActive('codeBlock')} />
          <ToolbarButton label="Left" onClick={() => editor.chain().focus().setTextAlign('left').run()} active={editor.isActive({ textAlign: 'left' })} />
          <ToolbarButton label="Center" onClick={() => editor.chain().focus().setTextAlign('center').run()} active={editor.isActive({ textAlign: 'center' })} />
          <ToolbarButton label="Right" onClick={() => editor.chain().focus().setTextAlign('right').run()} active={editor.isActive({ textAlign: 'right' })} />
          <ToolbarButton label="Justify" onClick={() => editor.chain().focus().setTextAlign('justify').run()} active={editor.isActive({ textAlign: 'justify' })} />
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <span className="mr-1 text-[10px] uppercase tracking-[0.16em] text-white/45">Layout</span>
          <ToolbarButton label="Insert Page Break" onClick={() => editor.chain().focus().setHorizontalRule().run()} />
          <ToolbarButton label={showPageBreaks ? 'Hide Page Guides' : 'Show Page Guides'} onClick={() => setShowPageBreaks((current) => !current)} active={showPageBreaks} />
          <ToolbarButton label="Link" onClick={openLinkEditor} active={editor.isActive('link') || linkEditorOpen} />
          <ToolbarButton label="Unlink" onClick={() => editor.chain().focus().unsetLink().run()} />
        </div>

        {linkEditorOpen && (
          <div className="flex flex-wrap items-center gap-2 border-t border-white/10 pt-2">
            <input
              value={linkDraft}
              onChange={(event) => setLinkDraft(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') {
                  event.preventDefault();
                  applyLink();
                }
                if (event.key === 'Escape') {
                  setLinkEditorOpen(false);
                }
              }}
              autoFocus
              placeholder="https://example.com"
              className="min-w-[220px] flex-1 rounded-full border border-white/15 bg-white/5 px-3 py-1.5 text-xs text-white outline-none placeholder:text-white/40"
            />
            <ToolbarButton label="Apply link" onClick={applyLink} />
            <ToolbarButton label="Cancel" onClick={() => setLinkEditorOpen(false)} />
          </div>
        )}
          </>
        )}
      </div>
      <div className="flex-1 min-h-0 overflow-auto">
        <div className={`docx-page-frame ${showPageBreaks ? 'show-page-breaks' : ''}`}>
          <EditorContent editor={editor} />
        </div>
      </div>
    </div>
  );
}

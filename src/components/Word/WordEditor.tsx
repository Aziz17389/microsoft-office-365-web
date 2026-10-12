import React, { useRef, useEffect, useState, useCallback } from 'react';
import { FormattingState, Language, WordEditorState } from '../../types/office';
import { getTranslation } from '../../i18n/translations';
import { ZoomIn, ZoomOut } from 'lucide-react';

interface WordEditorProps {
  lang: Language;
  state: WordEditorState;
  onChangeContent: (html: string) => void;
  onUpdateFormatting: (formatting: FormattingState) => void;
  isFormatPainterActive: boolean;
  onDeactivateFormatPainter: () => void;
  savedFormatPainterStyle?: FormattingState | null;
  onSetZoom: (zoom: number) => void;
  editorRefProp?: React.RefObject<HTMLDivElement | null>;
}

export const WordEditor: React.FC<WordEditorProps> = ({
  lang,
  state,
  onChangeContent,
  onUpdateFormatting,
  isFormatPainterActive,
  onDeactivateFormatPainter,
  savedFormatPainterStyle,
  onSetZoom,
  editorRefProp,
}) => {
  const localRef = useRef<HTMLDivElement>(null);
  const editorRef = editorRefProp || localRef;
  const isUpdatingFromState = useRef(false);

  const [wordCount, setWordCount] = useState(0);
  const [charCount, setCharCount] = useState(0);

  const t = getTranslation(lang);

  // Synchronize initial or loaded content without clobbering active user cursor
  useEffect(() => {
    if (editorRef.current && editorRef.current.innerHTML !== state.htmlContent) {
      if (document.activeElement !== editorRef.current) {
        editorRef.current.innerHTML = state.htmlContent;
        updateStats();
      }
    }
  }, [state.htmlContent]);

  const updateStats = useCallback(() => {
    if (!editorRef.current) return;
    const text = editorRef.current.innerText || '';
    const cleanWords = text.trim() ? text.trim().split(/\s+/).length : 0;
    setWordCount(cleanWords);
    setCharCount(text.length);
  }, [editorRef]);

  // Track selection state to update Ribbon toolbar formatting buttons
  const queryFormatting = useCallback(() => {
    if (!editorRef.current) return;

    try {
      const bold = document.queryCommandState('bold');
      const italic = document.queryCommandState('italic');
      const underline = document.queryCommandState('underline');
      const strikethrough = document.queryCommandState('strikeThrough');
      const subscript = document.queryCommandState('subscript');
      const superscript = document.queryCommandState('superscript');

      let align: 'left' | 'center' | 'right' | 'justify' = 'left';
      if (document.queryCommandState('justifyCenter')) align = 'center';
      else if (document.queryCommandState('justifyRight')) align = 'right';
      else if (document.queryCommandState('justifyFull')) align = 'justify';

      let fontName = document.queryCommandValue('fontName') || 'Segoe UI';
      // Clean up quotes from fontName
      fontName = fontName.replace(/['"]/g, '');

      const fontSize = document.queryCommandValue('fontSize') || '11';
      const foreColor = document.queryCommandValue('foreColor') || '#000000';
      const hiliteColor = document.queryCommandValue('hiliteColor') || 'transparent';

      onUpdateFormatting({
        bold,
        italic,
        underline,
        strikethrough,
        subscript,
        superscript,
        fontName,
        fontSize,
        foreColor,
        hiliteColor,
        align,
      });
    } catch {
      // Ignore query errors in non-editable states
    }
  }, [editorRef, onUpdateFormatting]);

  // Listen to selectionchange and input
  useEffect(() => {
    const handleSelectionChange = () => {
      const sel = window.getSelection();
      if (!sel || !editorRef.current) return;
      if (editorRef.current.contains(sel.anchorNode)) {
        queryFormatting();
      }
    };

    document.addEventListener('selectionchange', handleSelectionChange);
    return () => document.removeEventListener('selectionchange', handleSelectionChange);
  }, [queryFormatting, editorRef]);

  const handleInput = () => {
    if (!editorRef.current) return;
    const html = editorRef.current.innerHTML;
    onChangeContent(html);
    updateStats();
    queryFormatting();
  };

  // Format painter application
  const handleMouseUp = () => {
    queryFormatting();
    if (isFormatPainterActive && savedFormatPainterStyle) {
      if (savedFormatPainterStyle.bold) document.execCommand('bold');
      if (savedFormatPainterStyle.italic) document.execCommand('italic');
      if (savedFormatPainterStyle.underline) document.execCommand('underline');
      if (savedFormatPainterStyle.fontName) document.execCommand('fontName', false, savedFormatPainterStyle.fontName);
      if (savedFormatPainterStyle.foreColor) document.execCommand('foreColor', false, savedFormatPainterStyle.foreColor);
      onDeactivateFormatPainter();
    }
  };

  // Margins styling
  const getMarginStyle = () => {
    switch (state.margins) {
      case 'narrow':
        return 'p-8'; // 0.5 in approx
      case 'wide':
        return 'p-20'; // 1.5 in approx
      case 'normal':
      default:
        return 'p-14'; // 1.0 in approx
    }
  };

  // Dimensions based on orientation & paper size
  const getPageDimensions = () => {
    const isLandscape = state.orientation === 'landscape';
    if (isLandscape) {
      return 'w-[1056px] min-h-[816px]';
    }
    // Portrait
    switch (state.paperSize) {
      case 'Letter':
        return 'w-[816px] min-h-[1056px]';
      case 'Legal':
        return 'w-[816px] min-h-[1344px]';
      case 'A4':
      default:
        return 'w-[794px] min-h-[1123px]';
    }
  };

  return (
    <div className="flex-1 flex flex-col h-full bg-[#e1dfdd] dark:bg-[#1b1a19] overflow-hidden select-text">
      {/* Optional Top Ruler */}
      {state.showRuler && (
        <div className="no-print bg-[#f3f2f1] dark:bg-[#252423] border-b border-gray-300 dark:border-[#3b3a39] h-6 flex items-center justify-center select-none overflow-hidden relative shadow-inner">
          <div className="w-[794px] h-full flex items-end justify-between px-14 text-[9px] text-gray-500 font-mono">
            {Array.from({ length: 9 }).map((_, i) => (
              <div key={i} className="flex flex-col items-center">
                <span className="text-[8px]">{i + 1}</span>
                <div className="h-2 w-[1px] bg-gray-400" />
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Workspace Area */}
      <div
        className="flex-1 overflow-auto p-6 md:p-12 flex justify-center items-start transition-all"
        style={{
          backgroundImage: state.showGridlines
            ? 'radial-gradient(circle, #00000015 1px, transparent 1px)'
            : 'none',
          backgroundSize: '20px 20px',
        }}
      >
        <div
          style={{
            transform: `scale(${state.zoom / 100})`,
            transformOrigin: 'top center',
            transition: 'transform 0.15s ease-out',
          }}
          className="print-page-container"
        >
          {/* A4 Realistic Page Container with Realistic Drop Shadow */}
          <div
            className={`word-page-sheet bg-white dark:bg-[#252423] text-gray-900 dark:text-gray-100 shadow-[0_10px_30px_rgba(0,0,0,0.18)] dark:shadow-[0_10px_30px_rgba(0,0,0,0.6)] border border-gray-300 dark:border-[#3b3a39] rounded-sm transition-all focus-within:ring-1 focus-within:ring-[#185abd] ${getPageDimensions()} ${getMarginStyle()}`}
          >
            {/* WYSIWYG Editable Core */}
            <div
              ref={editorRef}
              contentEditable
              suppressContentEditableWarning
              onInput={handleInput}
              onKeyUp={queryFormatting}
              onMouseUp={handleMouseUp}
              className="outline-none min-h-[900px] leading-relaxed text-sm focus:outline-none"
              style={{
                direction: lang === 'ar' ? 'rtl' : 'ltr',
              }}
            />
          </div>
        </div>
      </div>

      {/* Live Status Bar */}
      <div className="status-bar no-print bg-[#f3f2f1] dark:bg-[#201f1e] border-t border-gray-300 dark:border-[#3b3a39] px-4 py-1 flex items-center justify-between text-xs text-gray-600 dark:text-gray-400 select-none">
        <div className="flex items-center gap-4">
          <span>{t.pageStatus}</span>
          <span className="hidden sm:inline-block">
            {wordCount} {t.wordsCount}
          </span>
          <span className="hidden md:inline-block">
            {charCount} {t.charsCount}
          </span>
          <span className="hidden lg:inline-block font-medium">
            {lang === 'ar' ? 'العربية' : 'English (United States)'}
          </span>
        </div>

        {/* Zoom Slider and Controls */}
        <div className="flex items-center gap-2">
          <button
            onClick={() => onSetZoom(Math.max(50, state.zoom - 10))}
            className="p-0.5 rounded hover:bg-gray-200 dark:hover:bg-[#323130]"
            title="Zoom Out"
          >
            <ZoomOut className="w-3.5 h-3.5" />
          </button>

          <input
            type="range"
            min="50"
            max="200"
            step="5"
            value={state.zoom}
            onChange={(e) => onSetZoom(parseInt(e.target.value, 10))}
            className="w-24 h-1.5 bg-gray-300 dark:bg-gray-700 rounded-lg appearance-none cursor-pointer accent-[#185abd]"
          />

          <button
            onClick={() => onSetZoom(Math.min(200, state.zoom + 10))}
            className="p-0.5 rounded hover:bg-gray-200 dark:hover:bg-[#323130]"
            title="Zoom In"
          >
            <ZoomIn className="w-3.5 h-3.5" />
          </button>

          <span className="font-mono text-[11px] w-9 text-right">{state.zoom}%</span>
        </div>
      </div>
    </div>
  );
};

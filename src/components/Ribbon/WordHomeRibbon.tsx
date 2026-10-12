import React, { useState } from 'react';
import {
  Scissors,
  Copy,
  Paintbrush,
  Bold,
  Italic,
  Underline,
  Strikethrough,
  Subscript,
  Superscript,
  Highlighter,
  Baseline,
  RemoveFormatting,
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignJustify,
  List,
  ListOrdered,
  Indent,
  Outdent,
  Search,
  Check,
  ChevronDown,
  Lock,
} from 'lucide-react';
import { FormattingState, GatedFeature, Language, PlanTier } from '../../types/office';
import { getTranslation } from '../../i18n/translations';

interface WordHomeRibbonProps {
  lang: Language;
  formatting: FormattingState;
  plan: PlanTier;
  onTriggerGated: (feature: GatedFeature) => void;
  onCommand: (command: string, value?: any) => void;
  onOpenFindReplace: () => void;
  isFormatPainterActive: boolean;
  onToggleFormatPainter: () => void;
}

const FONT_FAMILIES = [
  'Segoe UI',
  'Calibri',
  'Arial',
  'Times New Roman',
  'Cairo',
  'Tahoma',
  'Courier New',
  'Georgia',
];

const FONT_SIZES = ['8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '28', '36', '48', '72'];

const HIGHLIGHT_COLORS = [
  'transparent',
  '#ffff00', // yellow
  '#00ff00', // bright green
  '#00ffff', // cyan
  '#ff00ff', // magenta
  '#ff8080', // light red
  '#ffa500', // orange
  '#e6e6e6', // light gray
];

const TEXT_COLORS = [
  '#000000',
  '#185abd', // Word blue
  '#c00000', // Dark red
  '#ed7d31', // Orange
  '#ffc000', // Gold
  '#70ad47', // Green
  '#7030a0', // Purple
  '#595959', // Gray
  '#ffffff', // White
];

export const WordHomeRibbon: React.FC<WordHomeRibbonProps> = ({
  lang,
  formatting,
  plan,
  onTriggerGated,
  onCommand,
  onOpenFindReplace,
  isFormatPainterActive,
  onToggleFormatPainter,
}) => {
  const [showHighlightPalette, setShowHighlightPalette] = useState(false);
  const [showColorPalette, setShowColorPalette] = useState(false);
  const [showLineSpacing, setShowLineSpacing] = useState(false);
  const t = getTranslation(lang);

  const applyStyle = (styleType: string) => {
    switch (styleType) {
      case 'normal':
        onCommand('formatBlock', '<p>');
        break;
      case 'h1':
        onCommand('formatBlock', '<h1>');
        break;
      case 'h2':
        onCommand('formatBlock', '<h2>');
        break;
      case 'h3':
        onCommand('formatBlock', '<h3>');
        break;
      case 'quote':
        onCommand('formatBlock', '<blockquote>');
        break;
    }
  };

  return (
    <div className="flex items-stretch h-24 overflow-x-auto select-none px-2 py-1 text-gray-700 dark:text-gray-200">
      {/* 1. CLIPBOARD GROUP */}
      <div className="flex flex-col justify-between px-2 border-r border-gray-200 dark:border-[#3b3a39] shrink-0">
        <div className="flex items-center gap-1 my-auto">
          {/* Paste */}
          <button
            onClick={() => onCommand('paste')}
            title={t.paste}
            className="flex flex-col items-center justify-center p-1.5 rounded hover:bg-gray-200 dark:hover:bg-[#323130] transition-colors cursor-pointer w-12"
          >
            <Scissors className="w-5 h-5 text-[#185abd] dark:text-blue-400 rotate-90" />
            <span className="text-[10px] mt-0.5">{t.paste.split(' ')[0]}</span>
          </button>

          <div className="flex flex-col gap-0.5">
            {/* Cut */}
            <button
              onClick={() => onCommand('cut')}
              title={t.cut}
              className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] flex items-center gap-1.5 text-xs transition-colors cursor-pointer"
            >
              <Scissors className="w-3.5 h-3.5" />
              <span className="text-[11px]">{t.cut.split(' ')[0]}</span>
            </button>
            {/* Copy */}
            <button
              onClick={() => onCommand('copy')}
              title={t.copy}
              className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] flex items-center gap-1.5 text-xs transition-colors cursor-pointer"
            >
              <Copy className="w-3.5 h-3.5" />
              <span className="text-[11px]">{t.copy.split(' ')[0]}</span>
            </button>
            {/* Format Painter */}
            <button
              onClick={onToggleFormatPainter}
              title={t.formatPainter}
              className={`p-1 rounded flex items-center gap-1.5 text-xs transition-colors cursor-pointer ${
                isFormatPainterActive
                  ? 'bg-[#185abd] text-white'
                  : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              <Paintbrush className="w-3.5 h-3.5" />
              <span className="text-[11px]">{t.formatPainter}</span>
            </button>
          </div>
        </div>
        <div className="text-[10px] text-center text-gray-400 dark:text-gray-500 font-medium">
          {t.groupClipboard}
        </div>
      </div>

      {/* 2. FONT GROUP */}
      <div className="flex flex-col justify-between px-2.5 border-r border-gray-200 dark:border-[#3b3a39] shrink-0">
        <div className="space-y-1 my-auto">
          {/* Top row: font family & size */}
          <div className="flex items-center gap-1">
            {/* Font Family (Gated) */}
            <div className="relative">
              <select
                value={formatting.fontName}
                onChange={(e) => {
                  if (plan === 'free') {
                    onTriggerGated('font_family');
                  } else {
                    onCommand('fontName', e.target.value);
                  }
                }}
                className={`h-6 px-1.5 text-xs rounded border border-gray-300 dark:border-[#484644] bg-white dark:bg-[#1b1a19] text-gray-800 dark:text-gray-200 outline-none w-28 cursor-pointer ${
                  plan === 'free' ? 'pr-5' : ''
                }`}
                title={plan === 'free' ? t.lockedFeatureTooltip : t.fontFamily}
              >
                {FONT_FAMILIES.map((f) => (
                  <option key={f} value={f} style={{ fontFamily: f }}>
                    {f}
                  </option>
                ))}
              </select>
              {plan === 'free' && (
                <span
                  onClick={() => onTriggerGated('font_family')}
                  className="absolute right-1 top-1/2 -translate-y-1/2 text-amber-500 cursor-pointer pointer-events-none"
                  title={t.lockedFeatureTooltip}
                >
                  <Lock className="w-2.5 h-2.5" />
                </span>
              )}
            </div>

            {/* Font Size (Gated) */}
            <div className="relative">
              <select
                value={formatting.fontSize}
                onChange={(e) => {
                  if (plan === 'free') {
                    onTriggerGated('text_size');
                  } else {
                    onCommand('fontSize', e.target.value);
                  }
                }}
                className={`h-6 px-1 text-xs rounded border border-gray-300 dark:border-[#484644] bg-white dark:bg-[#1b1a19] text-gray-800 dark:text-gray-200 outline-none w-14 cursor-pointer ${
                  plan === 'free' ? 'pr-4' : ''
                }`}
                title={plan === 'free' ? t.lockedFeatureTooltip : t.fontSize}
              >
                {FONT_SIZES.map((s) => (
                  <option key={s} value={s}>
                    {s} pt
                  </option>
                ))}
              </select>
              {plan === 'free' && (
                <span
                  onClick={() => onTriggerGated('text_size')}
                  className="absolute right-1 top-1/2 -translate-y-1/2 text-amber-500 cursor-pointer pointer-events-none"
                  title={t.lockedFeatureTooltip}
                >
                  <Lock className="w-2.5 h-2.5" />
                </span>
              )}
            </div>

            {/* Increase & Decrease Font (Gated) */}
            <button
              onClick={() => {
                if (plan === 'free') {
                  onTriggerGated('text_size');
                } else {
                  const current = parseInt(formatting.fontSize, 10) || 12;
                  onCommand('fontSize', String(Math.min(72, current + 2)));
                }
              }}
              title={plan === 'free' ? `${t.increaseFont} (${t.lockedFeatureTooltip})` : t.increaseFont}
              className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] font-bold text-xs relative cursor-pointer"
            >
              A<span className="text-[9px]">^</span>
              {plan === 'free' && (
                <span className="absolute -top-0.5 -right-0.5 text-amber-500 text-[8px]">
                  🔒
                </span>
              )}
            </button>
            <button
              onClick={() => {
                if (plan === 'free') {
                  onTriggerGated('text_size');
                } else {
                  const current = parseInt(formatting.fontSize, 10) || 12;
                  onCommand('fontSize', String(Math.max(8, current - 2)));
                }
              }}
              title={plan === 'free' ? `${t.decreaseFont} (${t.lockedFeatureTooltip})` : t.decreaseFont}
              className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] font-bold text-xs relative cursor-pointer"
            >
              A<span className="text-[9px]">v</span>
              {plan === 'free' && (
                <span className="absolute -top-0.5 -right-0.5 text-amber-500 text-[8px]">
                  🔒
                </span>
              )}
            </button>

            {/* Clear formatting */}
            <button
              onClick={() => onCommand('removeFormat')}
              title={t.clearFormatting}
              className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] text-gray-500 hover:text-red-600"
            >
              <RemoveFormatting className="w-3.5 h-3.5" />
            </button>
          </div>

          {/* Bottom row: B, I, U, Strike, Sub, Sup, Highlight, Color */}
          <div className="flex items-center gap-0.5">
            {/* Bold */}
            <button
              onClick={() => onCommand('bold')}
              title={t.bold}
              className={`p-1 rounded transition-colors ${
                formatting.bold
                  ? 'bg-[#185abd] text-white'
                  : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              <Bold className="w-3.5 h-3.5 font-bold" />
            </button>
            {/* Italic */}
            <button
              onClick={() => onCommand('italic')}
              title={t.italic}
              className={`p-1 rounded transition-colors ${
                formatting.italic
                  ? 'bg-[#185abd] text-white'
                  : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              <Italic className="w-3.5 h-3.5" />
            </button>
            {/* Underline */}
            <button
              onClick={() => onCommand('underline')}
              title={t.underline}
              className={`p-1 rounded transition-colors ${
                formatting.underline
                  ? 'bg-[#185abd] text-white'
                  : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              <Underline className="w-3.5 h-3.5" />
            </button>
            {/* Strikethrough */}
            <button
              onClick={() => onCommand('strikeThrough')}
              title={t.strikethrough}
              className={`p-1 rounded transition-colors ${
                formatting.strikethrough
                  ? 'bg-[#185abd] text-white'
                  : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              <Strikethrough className="w-3.5 h-3.5" />
            </button>
            {/* Subscript */}
            <button
              onClick={() => onCommand('subscript')}
              title={t.subscript}
              className={`p-1 rounded transition-colors ${
                formatting.subscript
                  ? 'bg-[#185abd] text-white'
                  : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              <Subscript className="w-3.5 h-3.5" />
            </button>
            {/* Superscript */}
            <button
              onClick={() => onCommand('superscript')}
              title={t.superscript}
              className={`p-1 rounded transition-colors ${
                formatting.superscript
                  ? 'bg-[#185abd] text-white'
                  : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              <Superscript className="w-3.5 h-3.5" />
            </button>

            {/* Highlight color picker */}
            <div className="relative">
              <button
                onClick={() => {
                  setShowHighlightPalette(!showHighlightPalette);
                  setShowColorPalette(false);
                }}
                title={t.textHighlightColor}
                className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] flex items-center gap-0.5"
              >
                <Highlighter className="w-3.5 h-3.5" />
                <ChevronDown className="w-2.5 h-2.5 opacity-60" />
              </button>
              {showHighlightPalette && (
                <div className="absolute top-full left-0 mt-1 p-2 bg-white dark:bg-[#252423] rounded shadow-lg border border-gray-200 dark:border-[#3b3a39] z-50 grid grid-cols-4 gap-1.5 w-32">
                  {HIGHLIGHT_COLORS.map((c) => (
                    <button
                      key={c}
                      onClick={() => {
                        onCommand('hiliteColor', c);
                        setShowHighlightPalette(false);
                      }}
                      className="w-5 h-5 rounded border border-gray-300 dark:border-gray-600 hover:scale-110 transition-transform"
                      style={{ backgroundColor: c === 'transparent' ? '#ffffff' : c }}
                      title={c === 'transparent' ? 'No Color' : c}
                    />
                  ))}
                </div>
              )}
            </div>

            {/* Font Color Picker */}
            <div className="relative">
              <button
                onClick={() => {
                  setShowColorPalette(!showColorPalette);
                  setShowHighlightPalette(false);
                }}
                title={t.fontColor}
                className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] flex items-center gap-0.5"
              >
                <Baseline className="w-3.5 h-3.5" />
                <ChevronDown className="w-2.5 h-2.5 opacity-60" />
              </button>
              {showColorPalette && (
                <div className="absolute top-full left-0 mt-1 p-2 bg-white dark:bg-[#252423] rounded shadow-lg border border-gray-200 dark:border-[#3b3a39] z-50 grid grid-cols-3 gap-1.5 w-28">
                  {TEXT_COLORS.map((c) => (
                    <button
                      key={c}
                      onClick={() => {
                        onCommand('foreColor', c);
                        setShowColorPalette(false);
                      }}
                      className="w-5 h-5 rounded border border-gray-300 dark:border-gray-600 hover:scale-110 transition-transform"
                      style={{ backgroundColor: c }}
                      title={c}
                    />
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>
        <div className="text-[10px] text-center text-gray-400 dark:text-gray-500 font-medium">
          {t.groupFont}
        </div>
      </div>

      {/* 3. PARAGRAPH GROUP */}
      <div className="flex flex-col justify-between px-2.5 border-r border-gray-200 dark:border-[#3b3a39] shrink-0">
        <div className="space-y-1 my-auto">
          {/* Top row: Bullets, Numbering, Indents, Text Direction */}
          <div className="flex items-center gap-1">
            <button
              onClick={() => onCommand('insertUnorderedList')}
              title={t.bulletList}
              className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130]"
            >
              <List className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => onCommand('insertOrderedList')}
              title={t.numberedList}
              className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130]"
            >
              <ListOrdered className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => onCommand('outdent')}
              title={t.decreaseIndent}
              className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130]"
            >
              <Outdent className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => onCommand('indent')}
              title={t.increaseIndent}
              className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130]"
            >
              <Indent className="w-3.5 h-3.5" />
            </button>

            {/* Line Spacing */}
            <div className="relative">
              <button
                onClick={() => setShowLineSpacing(!showLineSpacing)}
                title={t.lineSpacing}
                className="px-1.5 py-0.5 text-[11px] rounded hover:bg-gray-200 dark:hover:bg-[#323130] border border-gray-200 dark:border-gray-700 flex items-center gap-1"
              >
                <span>↕</span>
                <span className="text-[10px]">1.5</span>
              </button>
              {showLineSpacing && (
                <div className="absolute top-full left-0 mt-1 bg-white dark:bg-[#252423] rounded shadow-lg border border-gray-200 dark:border-[#3b3a39] py-1 z-50 w-20">
                  {['1.0', '1.15', '1.5', '2.0', '2.5'].map((sp) => (
                    <button
                      key={sp}
                      onClick={() => {
                        onCommand('lineSpacing', sp);
                        setShowLineSpacing(false);
                      }}
                      className="w-full text-left px-2.5 py-1 text-xs hover:bg-gray-100 dark:hover:bg-[#323130]"
                    >
                      {sp}
                    </button>
                  ))}
                </div>
              )}
            </div>
          </div>

          {/* Bottom row: Alignments */}
          <div className="flex items-center gap-0.5">
            <button
              onClick={() => onCommand('justifyLeft')}
              title={t.alignLeft}
              className={`p-1 rounded transition-colors ${
                formatting.align === 'left' ? 'bg-[#185abd] text-white' : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              <AlignLeft className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => onCommand('justifyCenter')}
              title={t.alignCenter}
              className={`p-1 rounded transition-colors ${
                formatting.align === 'center' ? 'bg-[#185abd] text-white' : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              <AlignCenter className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => onCommand('justifyRight')}
              title={t.alignRight}
              className={`p-1 rounded transition-colors ${
                formatting.align === 'right' ? 'bg-[#185abd] text-white' : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              <AlignRight className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => onCommand('justifyFull')}
              title={t.justify}
              className={`p-1 rounded transition-colors ${
                formatting.align === 'justify' ? 'bg-[#185abd] text-white' : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              <AlignJustify className="w-3.5 h-3.5" />
            </button>
          </div>
        </div>
        <div className="text-[10px] text-center text-gray-400 dark:text-gray-500 font-medium">
          {t.groupParagraph}
        </div>
      </div>

      {/* 4. STYLES GROUP */}
      <div className="flex flex-col justify-between px-2.5 border-r border-gray-200 dark:border-[#3b3a39] shrink-0">
        <div className="flex items-center gap-1.5 my-auto overflow-hidden">
          {/* Normal Style Tile */}
          <button
            onClick={() => applyStyle('normal')}
            className="flex flex-col items-start justify-center px-2 py-1 h-12 rounded border border-gray-200 dark:border-[#3b3a39] hover:border-[#185abd] bg-white dark:bg-[#1b1a19] text-left transition-colors cursor-pointer w-16"
          >
            <span className="text-xs font-normal">AaBbCc</span>
            <span className="text-[10px] text-gray-500 dark:text-gray-400 mt-0.5">{t.styleNormal}</span>
          </button>

          {/* Heading 1 Tile */}
          <button
            onClick={() => applyStyle('h1')}
            className="flex flex-col items-start justify-center px-2 py-1 h-12 rounded border border-gray-200 dark:border-[#3b3a39] hover:border-[#185abd] bg-white dark:bg-[#1b1a19] text-left transition-colors cursor-pointer w-16"
          >
            <span className="text-sm font-bold text-[#185abd] dark:text-blue-400">AaBb</span>
            <span className="text-[10px] text-gray-500 dark:text-gray-400 mt-0.5">{t.styleHeading1}</span>
          </button>

          {/* Heading 2 Tile */}
          <button
            onClick={() => applyStyle('h2')}
            className="flex flex-col items-start justify-center px-2 py-1 h-12 rounded border border-gray-200 dark:border-[#3b3a39] hover:border-[#185abd] bg-white dark:bg-[#1b1a19] text-left transition-colors cursor-pointer w-16"
          >
            <span className="text-xs font-semibold text-[#185abd] dark:text-blue-400">AaBb</span>
            <span className="text-[10px] text-gray-500 dark:text-gray-400 mt-0.5">{t.styleHeading2}</span>
          </button>

          {/* Quote Tile */}
          <button
            onClick={() => applyStyle('quote')}
            className="hidden sm:flex flex-col items-start justify-center px-2 py-1 h-12 rounded border border-gray-200 dark:border-[#3b3a39] hover:border-[#185abd] bg-white dark:bg-[#1b1a19] text-left transition-colors cursor-pointer w-16"
          >
            <span className="text-xs italic text-gray-600 dark:text-gray-300">“Quote”</span>
            <span className="text-[10px] text-gray-500 dark:text-gray-400 mt-0.5">{t.styleQuote}</span>
          </button>
        </div>
        <div className="text-[10px] text-center text-gray-400 dark:text-gray-500 font-medium">
          {t.groupStyles}
        </div>
      </div>

      {/* 5. EDITING GROUP */}
      <div className="flex flex-col justify-between px-2.5 shrink-0">
        <div className="flex flex-col justify-center gap-1 my-auto">
          <button
            onClick={onOpenFindReplace}
            title={t.findAndReplace}
            className="flex items-center gap-1.5 px-2 py-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] text-xs transition-colors cursor-pointer"
          >
            <Search className="w-3.5 h-3.5 text-[#185abd] dark:text-blue-400" />
            <span className="text-[11px]">{t.findAndReplace.split(' ')[0]}</span>
          </button>

          <button
            onClick={() => onCommand('selectAll')}
            title={t.selectAll}
            className="flex items-center gap-1.5 px-2 py-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] text-xs transition-colors cursor-pointer"
          >
            <Check className="w-3.5 h-3.5 text-gray-500" />
            <span className="text-[11px]">{t.selectAll}</span>
          </button>
        </div>
        <div className="text-[10px] text-center text-gray-400 dark:text-gray-500 font-medium">
          {t.groupEditing}
        </div>
      </div>
    </div>
  );
};

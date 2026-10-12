import React, { useState } from 'react';
import {
  Scissors,
  Copy,
  Clipboard,
  Paintbrush,
  Bold,
  Italic,
  Underline,
  AlignLeft,
  AlignCenter,
  AlignRight,
  DollarSign,
  Percent,
  Search,
  Sigma,
  ArrowDownAZ,
  ArrowUpAZ,
  Trash2,
  ChevronDown,
  PaintBucket,
  Baseline,
  Lock,
} from 'lucide-react';
import { CellData, CellFormatType, GatedFeature, Language, PlanTier } from '../../types/office';
import { getTranslation } from '../../i18n/translations';

interface ExcelHomeRibbonProps {
  lang: Language;
  plan: PlanTier;
  onTriggerGated: (feature: GatedFeature) => void;
  activeCellData?: CellData;
  onUpdateFormat: (updates: Partial<CellData>) => void;
  onAutoSum: () => void;
  onSort: (ascending: boolean) => void;
  onClear: () => void;
  onOpenFindReplace: () => void;
  onOpenFormulaHelper: () => void;
}

const FONT_FAMILIES = ['Calibri', 'Segoe UI', 'Arial', 'Cairo', 'Tahoma', 'Times New Roman'];
const FONT_SIZES = [9, 10, 11, 12, 14, 16, 18, 20, 24];

const FILL_COLORS = [
  'transparent',
  '#ffff00', // Yellow
  '#d9ead3', // Soft green
  '#cfe2f3', // Soft blue
  '#fce5cd', // Soft orange
  '#f4cccc', // Soft red
  '#e1dfdd', // Office light gray
  '#ffffff', // White
];

const TEXT_COLORS = [
  '#000000',
  '#107c41', // Excel green
  '#c00000', // Dark red
  '#185abd', // Blue
  '#7030a0', // Purple
  '#ffffff', // White
];

export const ExcelHomeRibbon: React.FC<ExcelHomeRibbonProps> = ({
  lang,
  plan,
  onTriggerGated,
  activeCellData = { raw: '' },
  onUpdateFormat,
  onAutoSum,
  onSort,
  onClear,
  onOpenFindReplace,
  onOpenFormulaHelper,
}) => {
  const [showFillPalette, setShowFillPalette] = useState(false);
  const [showTextPalette, setShowTextPalette] = useState(false);
  const t = getTranslation(lang);

  const currentFormat = activeCellData.format || 'general';
  const isBold = Boolean(activeCellData.bold);
  const isItalic = Boolean(activeCellData.italic);
  const isUnderline = Boolean(activeCellData.underline);
  const currentAlign = activeCellData.align || 'left';
  const decimals = activeCellData.decimals !== undefined ? activeCellData.decimals : 2;

  return (
    <div className="flex items-stretch h-24 overflow-x-auto select-none px-2 py-1 text-gray-700 dark:text-gray-200">
      {/* 1. CLIPBOARD GROUP */}
      <div className="flex flex-col justify-between px-2 border-r border-gray-200 dark:border-[#3b3a39] shrink-0">
        <div className="flex items-center gap-1 my-auto">
          <button
            title={t.paste}
            className="flex flex-col items-center justify-center p-1.5 rounded hover:bg-gray-200 dark:hover:bg-[#323130] transition-colors cursor-pointer w-12"
          >
            <Clipboard className="w-5 h-5 text-[#107c41] dark:text-green-400" />
            <span className="text-[10px] mt-0.5">{t.paste.split(' ')[0]}</span>
          </button>

          <div className="flex flex-col gap-0.5">
            <button
              title={t.cut}
              className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] flex items-center gap-1.5 text-xs transition-colors cursor-pointer"
            >
              <Scissors className="w-3.5 h-3.5" />
              <span className="text-[11px]">{t.cut.split(' ')[0]}</span>
            </button>
            <button
              title={t.copy}
              className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] flex items-center gap-1.5 text-xs transition-colors cursor-pointer"
            >
              <Copy className="w-3.5 h-3.5" />
              <span className="text-[11px]">{t.copy.split(' ')[0]}</span>
            </button>
            <button
              title={t.formatPainter}
              className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] flex items-center gap-1.5 text-xs transition-colors cursor-pointer"
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
          {/* Top: family & size */}
          <div className="flex items-center gap-1">
            <div className="relative">
              <select
                onChange={(e) => {
                  if (plan === 'free') {
                    onTriggerGated('font_family');
                  } else {
                    onUpdateFormat({ fontFamily: e.target.value });
                  }
                }}
                className={`h-6 px-1.5 text-xs rounded border border-gray-300 dark:border-[#484644] bg-white dark:bg-[#1b1a19] text-gray-800 dark:text-gray-200 outline-none w-24 cursor-pointer ${
                  plan === 'free' ? 'pr-5' : ''
                }`}
                title={plan === 'free' ? t.lockedFeatureTooltip : t.fontFamily}
              >
                {FONT_FAMILIES.map((f) => (
                  <option key={f} value={f}>
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

            <div className="relative">
              <select
                onChange={(e) => {
                  if (plan === 'free') {
                    onTriggerGated('text_size');
                  } else {
                    onUpdateFormat({ fontSize: parseInt(e.target.value, 10) });
                  }
                }}
                className={`h-6 px-1 text-xs rounded border border-gray-300 dark:border-[#484644] bg-white dark:bg-[#1b1a19] text-gray-800 dark:text-gray-200 outline-none w-14 cursor-pointer ${
                  plan === 'free' ? 'pr-4' : ''
                }`}
                title={plan === 'free' ? t.lockedFeatureTooltip : t.fontSize}
              >
                {FONT_SIZES.map((s) => (
                  <option key={s} value={s}>
                    {s}
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
                  const cur = activeCellData.fontSize || 11;
                  onUpdateFormat({ fontSize: Math.min(72, cur + 2) });
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
                  const cur = activeCellData.fontSize || 11;
                  onUpdateFormat({ fontSize: Math.max(8, cur - 2) });
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
          </div>

          {/* Bottom: B, I, U, Fill, Font Color */}
          <div className="flex items-center gap-0.5">
            <button
              onClick={() => onUpdateFormat({ bold: !isBold })}
              title={t.bold}
              className={`p-1 rounded transition-colors ${
                isBold ? 'bg-[#107c41] text-white' : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              <Bold className="w-3.5 h-3.5 font-bold" />
            </button>

            <button
              onClick={() => onUpdateFormat({ italic: !isItalic })}
              title={t.italic}
              className={`p-1 rounded transition-colors ${
                isItalic ? 'bg-[#107c41] text-white' : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              <Italic className="w-3.5 h-3.5" />
            </button>

            <button
              onClick={() => onUpdateFormat({ underline: !isUnderline })}
              title={t.underline}
              className={`p-1 rounded transition-colors ${
                isUnderline ? 'bg-[#107c41] text-white' : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              <Underline className="w-3.5 h-3.5" />
            </button>

            {/* Fill Color */}
            <div className="relative">
              <button
                onClick={() => {
                  setShowFillPalette(!showFillPalette);
                  setShowTextPalette(false);
                }}
                title="Fill Color"
                className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] flex items-center gap-0.5"
              >
                <PaintBucket className="w-3.5 h-3.5" />
                <ChevronDown className="w-2.5 h-2.5 opacity-60" />
              </button>
              {showFillPalette && (
                <div className="absolute top-full left-0 mt-1 p-2 bg-white dark:bg-[#252423] rounded shadow-lg border border-gray-200 dark:border-[#3b3a39] z-50 grid grid-cols-4 gap-1.5 w-32">
                  {FILL_COLORS.map((c) => (
                    <button
                      key={c}
                      onClick={() => {
                        onUpdateFormat({ bgColor: c === 'transparent' ? undefined : c });
                        setShowFillPalette(false);
                      }}
                      className="w-5 h-5 rounded border border-gray-300 dark:border-gray-600 hover:scale-110 transition-transform"
                      style={{ backgroundColor: c === 'transparent' ? '#ffffff' : c }}
                      title={c}
                    />
                  ))}
                </div>
              )}
            </div>

            {/* Text Color */}
            <div className="relative">
              <button
                onClick={() => {
                  setShowTextPalette(!showTextPalette);
                  setShowFillPalette(false);
                }}
                title={t.fontColor}
                className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] flex items-center gap-0.5"
              >
                <Baseline className="w-3.5 h-3.5" />
                <ChevronDown className="w-2.5 h-2.5 opacity-60" />
              </button>
              {showTextPalette && (
                <div className="absolute top-full left-0 mt-1 p-2 bg-white dark:bg-[#252423] rounded shadow-lg border border-gray-200 dark:border-[#3b3a39] z-50 grid grid-cols-3 gap-1.5 w-28">
                  {TEXT_COLORS.map((c) => (
                    <button
                      key={c}
                      onClick={() => {
                        onUpdateFormat({ color: c });
                        setShowTextPalette(false);
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

      {/* 3. ALIGNMENT GROUP */}
      <div className="flex flex-col justify-between px-2.5 border-r border-gray-200 dark:border-[#3b3a39] shrink-0">
        <div className="space-y-1 my-auto">
          {/* Horizontal Alignment */}
          <div className="flex items-center gap-1">
            <button
              onClick={() => onUpdateFormat({ align: 'left' })}
              title={t.alignLeft}
              className={`p-1 rounded transition-colors ${
                currentAlign === 'left' ? 'bg-[#107c41] text-white' : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              <AlignLeft className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => onUpdateFormat({ align: 'center' })}
              title={t.alignCenter}
              className={`p-1 rounded transition-colors ${
                currentAlign === 'center' ? 'bg-[#107c41] text-white' : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              <AlignCenter className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => onUpdateFormat({ align: 'right' })}
              title={t.alignRight}
              className={`p-1 rounded transition-colors ${
                currentAlign === 'right' ? 'bg-[#107c41] text-white' : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              <AlignRight className="w-3.5 h-3.5" />
            </button>
          </div>

          {/* Merge & Center Button */}
          <button
            onClick={() => {
              onUpdateFormat({ align: 'center', bold: true });
            }}
            title={t.mergeAndCenter}
            className="w-full text-center py-0.5 px-1.5 text-[10px] rounded border border-gray-200 dark:border-gray-700 hover:bg-gray-200 dark:hover:bg-[#323130] transition-colors"
          >
            {t.mergeAndCenter}
          </button>
        </div>
        <div className="text-[10px] text-center text-gray-400 dark:text-gray-500 font-medium">
          {t.groupAlignment}
        </div>
      </div>

      {/* 4. NUMBER GROUP */}
      <div className="flex flex-col justify-between px-2.5 border-r border-gray-200 dark:border-[#3b3a39] shrink-0">
        <div className="space-y-1 my-auto">
          {/* Format selector */}
          <select
            value={currentFormat}
            onChange={(e) => onUpdateFormat({ format: e.target.value as CellFormatType })}
            className="h-6 px-1.5 text-xs rounded border border-gray-300 dark:border-[#484644] bg-white dark:bg-[#1b1a19] text-gray-800 dark:text-gray-200 outline-none w-28 cursor-pointer"
          >
            <option value="general">General</option>
            <option value="number">Number (1,234.00)</option>
            <option value="currency">Currency ($ / ر.س)</option>
            <option value="percent">Percentage (%)</option>
          </select>

          {/* Quick buttons: $, %, ,, .00, .0 */}
          <div className="flex items-center gap-1">
            <button
              onClick={() => onUpdateFormat({ format: 'currency', align: 'right' })}
              title={t.formatCurrency}
              className={`p-1 rounded text-xs transition-colors ${
                currentFormat === 'currency' ? 'bg-[#107c41] text-white' : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              <DollarSign className="w-3.5 h-3.5" />
            </button>

            <button
              onClick={() => onUpdateFormat({ format: 'percent', align: 'right' })}
              title={t.formatPercent}
              className={`p-1 rounded text-xs transition-colors ${
                currentFormat === 'percent' ? 'bg-[#107c41] text-white' : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              <Percent className="w-3.5 h-3.5" />
            </button>

            <button
              onClick={() => onUpdateFormat({ format: 'number', align: 'right' })}
              title={t.formatComma}
              className={`px-1.5 py-0.5 rounded text-xs font-bold font-mono transition-colors ${
                currentFormat === 'number' ? 'bg-[#107c41] text-white' : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              ,
            </button>

            <button
              onClick={() => onUpdateFormat({ decimals: Math.min(6, decimals + 1) })}
              title={t.increaseDecimal}
              className="px-1 py-0.5 rounded text-[10px] font-mono hover:bg-gray-200 dark:hover:bg-[#323130]"
            >
              .00→
            </button>

            <button
              onClick={() => onUpdateFormat({ decimals: Math.max(0, decimals - 1) })}
              title={t.decreaseDecimal}
              className="px-1 py-0.5 rounded text-[10px] font-mono hover:bg-gray-200 dark:hover:bg-[#323130]"
            >
              ←.0
            </button>
          </div>
        </div>
        <div className="text-[10px] text-center text-gray-400 dark:text-gray-500 font-medium">
          {t.groupNumber}
        </div>
      </div>

      {/* 5. EDITING / CALCULATION GROUP */}
      <div className="flex flex-col justify-between px-2.5 shrink-0">
        <div className="flex items-center gap-1.5 my-auto">
          {/* AutoSum */}
          <button
            onClick={onAutoSum}
            title={t.autoSum}
            className="flex flex-col items-center justify-center p-1.5 rounded hover:bg-gray-200 dark:hover:bg-[#323130] transition-colors cursor-pointer w-14"
          >
            <Sigma className="w-5 h-5 text-[#107c41] dark:text-green-400" />
            <span className="text-[10px] mt-0.5 font-medium">{t.autoSum.split(' ')[0]}</span>
          </button>

          <div className="flex flex-col gap-0.5">
            {/* Sort A-Z */}
            <button
              onClick={() => onSort(true)}
              title={t.sortAsc}
              className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] flex items-center gap-1.5 text-xs transition-colors cursor-pointer"
            >
              <ArrowDownAZ className="w-3.5 h-3.5" />
              <span className="text-[11px]">{t.sortAsc}</span>
            </button>
            {/* Sort Z-A */}
            <button
              onClick={() => onSort(false)}
              title={t.sortDesc}
              className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] flex items-center gap-1.5 text-xs transition-colors cursor-pointer"
            >
              <ArrowUpAZ className="w-3.5 h-3.5" />
              <span className="text-[11px]">{t.sortDesc}</span>
            </button>
            {/* Clear */}
            <button
              onClick={onClear}
              title="Clear contents"
              className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] text-red-600 flex items-center gap-1.5 text-xs transition-colors cursor-pointer"
            >
              <Trash2 className="w-3.5 h-3.5" />
              <span className="text-[11px]">Clear</span>
            </button>
          </div>

          <div className="flex flex-col gap-0.5 pl-1 border-l border-gray-200 dark:border-[#3b3a39]">
            <button
              onClick={onOpenFindReplace}
              title={t.findAndReplace}
              className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] flex items-center gap-1.5 text-xs transition-colors cursor-pointer"
            >
              <Search className="w-3.5 h-3.5" />
              <span className="text-[11px]">{t.find}</span>
            </button>
            <button
              onClick={onOpenFormulaHelper}
              title={t.fxButton}
              className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] text-[#107c41] font-bold font-serif flex items-center gap-1.5 text-xs transition-colors cursor-pointer"
            >
              <span className="italic">fx</span>
              <span className="text-[11px] font-sans font-normal">{t.fxButton}</span>
            </button>
          </div>
        </div>
        <div className="text-[10px] text-center text-gray-400 dark:text-gray-500 font-medium">
          {t.groupEditing}
        </div>
      </div>
    </div>
  );
};

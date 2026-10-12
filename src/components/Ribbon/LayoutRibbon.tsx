import React from 'react';
import {
  FileText,
  RotateCw,
  Maximize,
  LayoutTemplate,
  Lock,
} from 'lucide-react';
import {
  GatedFeature,
  Language,
  PlanTier,
  WordPageMargins,
  WordPageOrientation,
  WordPaperSize,
} from '../../types/office';
import { getTranslation } from '../../i18n/translations';

interface LayoutRibbonProps {
  lang: Language;
  plan: PlanTier;
  onTriggerGated: (feature: GatedFeature) => void;
  margins: WordPageMargins;
  orientation: WordPageOrientation;
  paperSize: WordPaperSize;
  onChangeMargins: (m: WordPageMargins) => void;
  onChangeOrientation: (o: WordPageOrientation) => void;
  onChangePaperSize: (s: WordPaperSize) => void;
}

export const LayoutRibbon: React.FC<LayoutRibbonProps> = ({
  lang,
  plan,
  onTriggerGated,
  margins,
  orientation,
  paperSize,
  onChangeMargins,
  onChangeOrientation,
  onChangePaperSize,
}) => {
  const t = getTranslation(lang);

  return (
    <div className="flex items-stretch h-24 overflow-x-auto select-none px-2 py-1 text-gray-700 dark:text-gray-200">
      {/* 1. MARGINS */}
      <div className="flex flex-col justify-between px-3 border-r border-gray-200 dark:border-[#3b3a39] shrink-0">
        <div className="flex items-center gap-1.5 my-auto">
          <div className="flex flex-col items-center">
            <LayoutTemplate className="w-5 h-5 text-[#185abd] dark:text-blue-400 mb-1" />
            <span className="text-[10px] text-gray-500 dark:text-gray-400">{t.margins}</span>
          </div>
          <div className="flex flex-col gap-1">
            <button
              onClick={() => onChangeMargins('normal')}
              className={`px-2 py-0.5 text-xs rounded transition-colors text-left ${
                margins === 'normal'
                  ? 'bg-[#185abd] text-white font-medium'
                  : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              {t.marginNormal}
            </button>
            <button
              onClick={() => onChangeMargins('narrow')}
              className={`px-2 py-0.5 text-xs rounded transition-colors text-left ${
                margins === 'narrow'
                  ? 'bg-[#185abd] text-white font-medium'
                  : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              {t.marginNarrow}
            </button>
            <button
              onClick={() => onChangeMargins('wide')}
              className={`px-2 py-0.5 text-xs rounded transition-colors text-left ${
                margins === 'wide'
                  ? 'bg-[#185abd] text-white font-medium'
                  : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
              }`}
            >
              {t.marginWide}
            </button>
          </div>
        </div>
        <div className="text-[10px] text-center text-gray-400 dark:text-gray-500 font-medium">
          {t.margins}
        </div>
      </div>

      {/* 2. ORIENTATION */}
      <div className="flex flex-col justify-between px-3 border-r border-gray-200 dark:border-[#3b3a39] shrink-0">
        <div className="flex items-center gap-2 my-auto">
          <button
            onClick={() => onChangeOrientation('portrait')}
            className={`flex flex-col items-center justify-center p-2 rounded transition-colors w-16 ${
              orientation === 'portrait'
                ? 'bg-[#185abd] text-white'
                : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
            }`}
          >
            <FileText className="w-5 h-5" />
            <span className="text-[11px] mt-1 font-medium">{t.portrait}</span>
          </button>

          <button
            onClick={() => onChangeOrientation('landscape')}
            className={`flex flex-col items-center justify-center p-2 rounded transition-colors w-16 ${
              orientation === 'landscape'
                ? 'bg-[#185abd] text-white'
                : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
            }`}
          >
            <RotateCw className="w-5 h-5" />
            <span className="text-[11px] mt-1 font-medium">{t.landscape}</span>
          </button>
        </div>
        <div className="text-[10px] text-center text-gray-400 dark:text-gray-500 font-medium">
          {t.orientation}
        </div>
      </div>

      {/* 3. PAPER SIZE (Gated behind Office 365+) */}
      <div className="flex flex-col justify-between px-3 shrink-0">
        <div className="flex items-center gap-2 my-auto">
          <div className="flex flex-col items-center">
            <Maximize className="w-5 h-5 text-gray-600 dark:text-gray-300 mb-1" />
            <span className="text-[10px] text-gray-500 dark:text-gray-400 flex items-center gap-1">
              <span>{t.paperSize}</span>
              {plan === 'free' && <Lock className="w-2.5 h-2.5 text-amber-500" />}
            </span>
          </div>

          <div className="flex flex-col gap-1">
            {(['A4', 'Letter', 'Legal'] as WordPaperSize[]).map((sz) => (
              <button
                key={sz}
                onClick={() => {
                  if (plan === 'free') {
                    onTriggerGated('page_size');
                  } else {
                    onChangePaperSize(sz);
                  }
                }}
                title={plan === 'free' ? `${sz} (${t.lockedFeatureTooltip})` : sz}
                className={`px-3 py-0.5 text-xs rounded transition-colors text-left font-mono flex items-center justify-between gap-2 ${
                  paperSize === sz
                    ? 'bg-[#185abd] text-white font-semibold'
                    : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
                }`}
              >
                <span>{sz}</span>
                {plan === 'free' && <Lock className="w-2.5 h-2.5 text-amber-500" />}
              </button>
            ))}
          </div>
        </div>
        <div className="text-[10px] text-center text-gray-400 dark:text-gray-500 font-medium">
          {t.groupPageSetup}
        </div>
      </div>
    </div>
  );
};

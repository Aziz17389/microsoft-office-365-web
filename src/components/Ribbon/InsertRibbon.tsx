import React, { useRef } from 'react';
import {
  Table,
  Image as ImageIcon,
  Link as LinkIcon,
  Minus,
  Split,
  Smile,
  Hash,
  Lock,
} from 'lucide-react';
import { AppMode, GatedFeature, Language, PlanTier } from '../../types/office';
import { getTranslation } from '../../i18n/translations';

interface InsertRibbonProps {
  appMode: AppMode;
  lang: Language;
  plan: PlanTier;
  onTriggerGated: (feature: GatedFeature) => void;
  onOpenTableModal: () => void;
  onOpenLinkModal: () => void;
  onOpenSymbolModal: () => void;
  onInsertImage: (dataUrl: string) => void;
  onInsertHorizontalRule?: () => void;
  onInsertPageBreak?: () => void;
}

export const InsertRibbon: React.FC<InsertRibbonProps> = ({
  appMode,
  lang,
  plan,
  onTriggerGated,
  onOpenTableModal,
  onOpenLinkModal,
  onOpenSymbolModal,
  onInsertImage,
  onInsertHorizontalRule,
  onInsertPageBreak,
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const t = getTranslation(lang);
  const isWord = appMode === 'word';
  const brandColor = isWord ? 'text-[#185abd] dark:text-blue-400' : 'text-[#107c41] dark:text-green-400';

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (loadEvent) => {
        if (loadEvent.target?.result) {
          onInsertImage(loadEvent.target.result as string);
        }
      };
      reader.readAsDataURL(file);
    }
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  return (
    <div className="flex items-stretch h-24 overflow-x-auto select-none px-2 py-1 text-gray-700 dark:text-gray-200">
      {/* 1. TABLES GROUP (Gated behind Office 365+) */}
      <div className="flex flex-col justify-between px-3 border-r border-gray-200 dark:border-[#3b3a39] shrink-0">
        <div className="flex items-center my-auto">
          <button
            onClick={() => {
              if (plan === 'free') {
                onTriggerGated('tables');
              } else {
                onOpenTableModal();
              }
            }}
            title={plan === 'free' ? `${t.insertTable} (${t.lockedFeatureTooltip})` : t.insertTable}
            className="flex flex-col items-center justify-center p-2 rounded hover:bg-gray-200 dark:hover:bg-[#323130] transition-colors cursor-pointer w-14 relative group"
          >
            <Table className={`w-6 h-6 ${brandColor}`} />
            <span className="text-[11px] mt-1 font-medium">{t.insertTable}</span>
            {plan === 'free' && (
              <span className="absolute top-1 right-1 rtl:right-auto rtl:left-1 bg-amber-400 text-black text-[9px] rounded-full p-0.5 font-bold shadow-xs">
                <Lock className="w-2.5 h-2.5" />
              </span>
            )}
          </button>
        </div>
        <div className="text-[10px] text-center text-gray-400 dark:text-gray-500 font-medium">
          {t.groupTables}
        </div>
      </div>

      {/* 2. ILLUSTRATIONS GROUP */}
      <div className="flex flex-col justify-between px-3 border-r border-gray-200 dark:border-[#3b3a39] shrink-0">
        <div className="flex items-center my-auto">
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileChange}
            accept="image/*"
            className="hidden"
          />
          <button
            onClick={() => fileInputRef.current?.click()}
            className="flex flex-col items-center justify-center p-2 rounded hover:bg-gray-200 dark:hover:bg-[#323130] transition-colors cursor-pointer w-14"
          >
            <ImageIcon className={`w-6 h-6 ${brandColor}`} />
            <span className="text-[11px] mt-1 font-medium">{t.insertPicture}</span>
          </button>
        </div>
        <div className="text-[10px] text-center text-gray-400 dark:text-gray-500 font-medium">
          {t.groupIllustrations}
        </div>
      </div>

      {/* 3. LINKS GROUP (Gated behind Office 365+) */}
      <div className="flex flex-col justify-between px-3 border-r border-gray-200 dark:border-[#3b3a39] shrink-0">
        <div className="flex items-center my-auto">
          <button
            onClick={() => {
              if (plan === 'free') {
                onTriggerGated('hyperlinks');
              } else {
                onOpenLinkModal();
              }
            }}
            title={plan === 'free' ? `${t.insertLink} (${t.lockedFeatureTooltip})` : t.insertLink}
            className="flex flex-col items-center justify-center p-2 rounded hover:bg-gray-200 dark:hover:bg-[#323130] transition-colors cursor-pointer w-14 relative group"
          >
            <LinkIcon className={`w-6 h-6 ${brandColor}`} />
            <span className="text-[11px] mt-1 font-medium">{t.insertLink.split(' ')[0]}</span>
            {plan === 'free' && (
              <span className="absolute top-1 right-1 rtl:right-auto rtl:left-1 bg-amber-400 text-black text-[9px] rounded-full p-0.5 font-bold shadow-xs">
                <Lock className="w-2.5 h-2.5" />
              </span>
            )}
          </button>
        </div>
        <div className="text-[10px] text-center text-gray-400 dark:text-gray-500 font-medium">
          {t.groupLinks}
        </div>
      </div>

      {/* 4. BREAKS & RULES (WORD ONLY) */}
      {isWord && (
        <div className="flex flex-col justify-between px-2.5 border-r border-gray-200 dark:border-[#3b3a39] shrink-0">
          <div className="flex items-center gap-1 my-auto">
            <button
              onClick={onInsertHorizontalRule}
              title={t.insertHorizontalLine}
              className="flex flex-col items-center justify-center p-1.5 rounded hover:bg-gray-200 dark:hover:bg-[#323130] text-xs transition-colors cursor-pointer w-16"
            >
              <Minus className="w-5 h-5 text-gray-600 dark:text-gray-300" />
              <span className="text-[10px] mt-0.5 truncate max-w-full">Divider</span>
            </button>

            <button
              onClick={onInsertPageBreak}
              title={t.insertPageBreak}
              className="flex flex-col items-center justify-center p-1.5 rounded hover:bg-gray-200 dark:hover:bg-[#323130] text-xs transition-colors cursor-pointer w-16"
            >
              <Split className="w-5 h-5 text-gray-600 dark:text-gray-300" />
              <span className="text-[10px] mt-0.5 truncate max-w-full">Page Break</span>
            </button>
          </div>
          <div className="text-[10px] text-center text-gray-400 dark:text-gray-500 font-medium">
            Breaks
          </div>
        </div>
      )}

      {/* 5. SYMBOLS & EMOJIS */}
      <div className="flex flex-col justify-between px-3 shrink-0">
        <div className="flex items-center gap-1.5 my-auto">
          <button
            onClick={onOpenSymbolModal}
            className="flex flex-col items-center justify-center p-2 rounded hover:bg-gray-200 dark:hover:bg-[#323130] transition-colors cursor-pointer w-16"
          >
            <div className="flex items-center gap-0.5">
              <Hash className="w-4 h-4 text-purple-600" />
              <Smile className="w-4 h-4 text-yellow-500" />
            </div>
            <span className="text-[11px] mt-1 font-medium">{t.insertSymbol.split(' ')[0]}</span>
          </button>
        </div>
        <div className="text-[10px] text-center text-gray-400 dark:text-gray-500 font-medium">
          Symbols
        </div>
      </div>
    </div>
  );
};

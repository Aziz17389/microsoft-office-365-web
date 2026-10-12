import React, { useState } from 'react';
import { X, Smile, Hash } from 'lucide-react';
import { Language } from '../../types/office';
import { getTranslation } from '../../i18n/translations';

interface InsertSymbolModalProps {
  isOpen: boolean;
  onClose: () => void;
  lang: Language;
  onInsertSymbol: (symbol: string) => void;
}

const SYMBOLS = [
  '©', '®', '™', '§', '¶', '°', '±', '×', '÷', '≠', '≈', '≤', '≥', '∞',
  '√', '∑', '∏', 'π', 'µ', '€', '$', '£', '¥', '₹', '₽', '₩', '₺', '﷼',
  '¢', '‰', '•', '–', '—', '…', '¿', '¡', '«', '»', '“', '”', '‘', '’',
  '←', '↑', '→', '↓', '↔', '↕', '⇒', '⇔', '✓', '✗', '✔', '★', '☆', '♥',
];

const EMOJIS = [
  '😀', '😊', '🚀', '💡', '🔥', '✨', '⭐', '🎉', '🏆', '🎯',
  '📊', '📈', '📉', '📋', '📁', '📂', '📄', '📅', '🕒', '📌',
  '💼', '🏢', '💻', '🖥️', '📱', '✉️', '📞', '🔍', '⚙️', '🔒',
  '✅', '❌', '⚠️', 'ℹ️', '💬', '📢', '👍', '👏', '🤝', '🙌',
];

export const InsertSymbolModal: React.FC<InsertSymbolModalProps> = ({
  isOpen,
  onClose,
  lang,
  onInsertSymbol,
}) => {
  const [tab, setTab] = useState<'symbols' | 'emojis'>('symbols');
  const t = getTranslation(lang);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/45 backdrop-blur-[1px] p-4">
      <div className="w-full max-w-md rounded-lg bg-white dark:bg-[#252423] text-gray-800 dark:text-gray-100 shadow-2xl border border-gray-300 dark:border-[#3b3a39] overflow-hidden">
        {/* Header */}
        <div className="flex items-center justify-between px-4 py-3 bg-gray-100 dark:bg-[#1f1e1d] border-b border-gray-200 dark:border-[#3b3a39]">
          <div className="flex items-center gap-2">
            <Hash className="w-4 h-4 text-[#185abd] dark:text-blue-400" />
            <h3 className="font-semibold text-sm">{t.insertSymbol}</h3>
          </div>
          <button onClick={onClose} className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130]">
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Tab switch */}
        <div className="flex border-b border-gray-200 dark:border-[#3b3a39] px-4 pt-2 gap-4">
          <button
            onClick={() => setTab('symbols')}
            className={`pb-2 text-xs font-medium border-b-2 transition-colors flex items-center gap-1.5 ${
              tab === 'symbols'
                ? 'border-[#185abd] text-[#185abd] dark:text-blue-400 dark:border-blue-400 font-semibold'
                : 'border-transparent text-gray-500 hover:text-gray-800 dark:hover:text-gray-200'
            }`}
          >
            <Hash className="w-3.5 h-3.5" />
            Symbols
          </button>
          <button
            onClick={() => setTab('emojis')}
            className={`pb-2 text-xs font-medium border-b-2 transition-colors flex items-center gap-1.5 ${
              tab === 'emojis'
                ? 'border-[#185abd] text-[#185abd] dark:text-blue-400 dark:border-blue-400 font-semibold'
                : 'border-transparent text-gray-500 hover:text-gray-800 dark:hover:text-gray-200'
            }`}
          >
            <Smile className="w-3.5 h-3.5" />
            Emojis
          </button>
        </div>

        {/* Grid */}
        <div className="p-4 max-h-64 overflow-y-auto">
          <div className="grid grid-cols-8 gap-2 text-center">
            {(tab === 'symbols' ? SYMBOLS : EMOJIS).map((item, idx) => (
              <button
                key={idx}
                onClick={() => {
                  onInsertSymbol(item);
                  onClose();
                }}
                className="w-9 h-9 text-lg rounded hover:bg-[#185abd]/10 dark:hover:bg-[#323130] hover:text-[#185abd] dark:hover:text-blue-400 transition-colors flex items-center justify-center border border-transparent hover:border-gray-300 dark:hover:border-gray-600"
              >
                {item}
              </button>
            ))}
          </div>
        </div>

        <div className="flex items-center justify-end px-4 py-2.5 bg-gray-50 dark:bg-[#1f1e1d] border-t border-gray-200 dark:border-[#3b3a39]">
          <button
            onClick={onClose}
            className="px-3 py-1.5 text-xs font-medium rounded hover:bg-gray-200 dark:hover:bg-[#323130] text-gray-600 dark:text-gray-300"
          >
            {t.close}
          </button>
        </div>
      </div>
    </div>
  );
};

import React, { useState } from 'react';
import { X, Search, Replace as ReplaceIcon } from 'lucide-react';
import { Language } from '../../types/office';
import { getTranslation } from '../../i18n/translations';

interface FindReplaceModalProps {
  isOpen: boolean;
  onClose: () => void;
  lang: Language;
  onFind: (search: string, matchCase: boolean) => void;
  onReplace: (search: string, replaceWith: string, matchCase: boolean) => void;
  onReplaceAll: (search: string, replaceWith: string, matchCase: boolean) => void;
}

export const FindReplaceModal: React.FC<FindReplaceModalProps> = ({
  isOpen,
  onClose,
  lang,
  onFind,
  onReplace,
  onReplaceAll,
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [replaceTerm, setReplaceTerm] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const t = getTranslation(lang);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/45 backdrop-blur-[1px] p-4 animate-in fade-in duration-150">
      <div className="w-full max-w-md rounded-lg bg-white dark:bg-[#252423] text-gray-800 dark:text-gray-100 shadow-2xl border border-gray-300 dark:border-[#3b3a39] overflow-hidden">
        {/* Header */}
        <div className="flex items-center justify-between px-4 py-3 bg-gray-100 dark:bg-[#1f1e1d] border-b border-gray-200 dark:border-[#3b3a39]">
          <div className="flex items-center gap-2">
            <Search className="w-4 h-4 text-[#185abd] dark:text-blue-400" />
            <h3 className="font-semibold text-sm">{t.findTitle}</h3>
          </div>
          <button
            onClick={onClose}
            className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Body */}
        <div className="p-4 space-y-4">
          <div>
            <label className="block text-xs font-medium text-gray-600 dark:text-gray-300 mb-1">
              {t.findWhat}
            </label>
            <input
              type="text"
              autoFocus
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full px-3 py-1.5 text-sm rounded border border-gray-300 dark:border-[#484644] bg-white dark:bg-[#1b1a19] focus:outline-none focus:ring-2 focus:ring-[#185abd]"
              placeholder="e.g. project"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-gray-600 dark:text-gray-300 mb-1">
              {t.replaceWith}
            </label>
            <input
              type="text"
              value={replaceTerm}
              onChange={(e) => setReplaceTerm(e.target.value)}
              className="w-full px-3 py-1.5 text-sm rounded border border-gray-300 dark:border-[#484644] bg-white dark:bg-[#1b1a19] focus:outline-none focus:ring-2 focus:ring-[#185abd]"
              placeholder="e.g. program"
            />
          </div>

          <div className="flex items-center gap-2 pt-1">
            <input
              type="checkbox"
              id="matchCase"
              checked={matchCase}
              onChange={(e) => setMatchCase(e.target.checked)}
              className="rounded border-gray-300 text-[#185abd] focus:ring-[#185abd]"
            />
            <label htmlFor="matchCase" className="text-xs text-gray-600 dark:text-gray-300 cursor-pointer">
              {t.matchCase}
            </label>
          </div>
        </div>

        {/* Footer */}
        <div className="flex items-center justify-end gap-2 px-4 py-3 bg-gray-50 dark:bg-[#1f1e1d] border-t border-gray-200 dark:border-[#3b3a39]">
          <button
            onClick={() => onFind(searchTerm, matchCase)}
            disabled={!searchTerm.trim()}
            className="px-3 py-1.5 text-xs font-medium rounded bg-gray-200 dark:bg-[#323130] hover:bg-gray-300 dark:hover:bg-[#3b3a39] disabled:opacity-50 transition-colors"
          >
            {t.findNext}
          </button>
          <button
            onClick={() => onReplace(searchTerm, replaceTerm, matchCase)}
            disabled={!searchTerm.trim()}
            className="px-3 py-1.5 text-xs font-medium rounded bg-gray-200 dark:bg-[#323130] hover:bg-gray-300 dark:hover:bg-[#3b3a39] disabled:opacity-50 transition-colors flex items-center gap-1.5"
          >
            <ReplaceIcon className="w-3.5 h-3.5" />
            {t.replace}
          </button>
          <button
            onClick={() => onReplaceAll(searchTerm, replaceTerm, matchCase)}
            disabled={!searchTerm.trim()}
            className="px-3 py-1.5 text-xs font-medium rounded bg-[#185abd] hover:bg-[#104a9e] text-white disabled:opacity-50 transition-colors"
          >
            {t.replaceAll}
          </button>
          <button
            onClick={onClose}
            className="px-3 py-1.5 text-xs font-medium rounded hover:bg-gray-200 dark:hover:bg-[#323130] text-gray-600 dark:text-gray-300 transition-colors"
          >
            {t.close}
          </button>
        </div>
      </div>
    </div>
  );
};

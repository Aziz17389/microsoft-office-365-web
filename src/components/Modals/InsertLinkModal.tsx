import React, { useState } from 'react';
import { X, Link as LinkIcon } from 'lucide-react';
import { Language } from '../../types/office';
import { getTranslation } from '../../i18n/translations';

interface InsertLinkModalProps {
  isOpen: boolean;
  onClose: () => void;
  lang: Language;
  initialText?: string;
  onInsertLink: (text: string, url: string) => void;
}

export const InsertLinkModal: React.FC<InsertLinkModalProps> = ({
  isOpen,
  onClose,
  lang,
  initialText = '',
  onInsertLink,
}) => {
  const [text, setText] = useState(initialText);
  const [url, setUrl] = useState('https://');
  const t = getTranslation(lang);

  if (!isOpen) return null;

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!url.trim()) return;
    onInsertLink(text.trim() || url.trim(), url.trim());
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/45 backdrop-blur-[1px] p-4">
      <div className="w-full max-w-sm rounded-lg bg-white dark:bg-[#252423] text-gray-800 dark:text-gray-100 shadow-2xl border border-gray-300 dark:border-[#3b3a39] overflow-hidden">
        <div className="flex items-center justify-between px-4 py-3 bg-gray-100 dark:bg-[#1f1e1d] border-b border-gray-200 dark:border-[#3b3a39]">
          <div className="flex items-center gap-2">
            <LinkIcon className="w-4 h-4 text-[#185abd] dark:text-blue-400" />
            <h3 className="font-semibold text-sm">{t.insertLink}</h3>
          </div>
          <button onClick={onClose} className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130]">
            <X className="w-4 h-4" />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-4 space-y-3">
          <div>
            <label className="block text-xs font-medium text-gray-600 dark:text-gray-300 mb-1">
              {t.linkText}
            </label>
            <input
              type="text"
              value={text}
              onChange={(e) => setText(e.target.value)}
              className="w-full px-3 py-1.5 text-sm rounded border border-gray-300 dark:border-[#484644] bg-white dark:bg-[#1b1a19] focus:outline-none focus:ring-2 focus:ring-[#185abd]"
              placeholder="e.g. Official Documentation"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-gray-600 dark:text-gray-300 mb-1">
              {t.linkUrl}
            </label>
            <input
              type="url"
              required
              value={url}
              onChange={(e) => setUrl(e.target.value)}
              className="w-full px-3 py-1.5 text-sm rounded border border-gray-300 dark:border-[#484644] bg-white dark:bg-[#1b1a19] focus:outline-none focus:ring-2 focus:ring-[#185abd]"
              placeholder="https://example.com"
            />
          </div>

          <div className="flex items-center justify-end gap-2 pt-3 border-t border-gray-200 dark:border-[#3b3a39]">
            <button
              type="button"
              onClick={onClose}
              className="px-3 py-1.5 text-xs font-medium rounded hover:bg-gray-200 dark:hover:bg-[#323130] text-gray-600 dark:text-gray-300"
            >
              {t.cancel}
            </button>
            <button
              type="submit"
              className="px-4 py-1.5 text-xs font-medium rounded bg-[#185abd] hover:bg-[#104a9e] text-white"
            >
              {t.insert}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};

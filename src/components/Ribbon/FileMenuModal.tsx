import React, { useState } from 'react';
import {
  ArrowLeft,
  FilePlus,
  Save,
  Download,
  Printer,
  Info,
  LogOut,
  FileText,
  FileSpreadsheet,
  CheckCircle2,
  Copy,
  Lock,
  Crown,
} from 'lucide-react';
import { AppMode, GatedFeature, Language, PlanTier } from '../../types/office';
import { getTranslation } from '../../i18n/translations';

interface FileMenuModalProps {
  isOpen: boolean;
  onClose: () => void;
  appMode: AppMode;
  lang: Language;
  plan: PlanTier;
  onTriggerGated: (feature: GatedFeature) => void;
  docTitle: string;
  onNew: () => void;
  onSave: () => void;
  onSaveAs: (newTitle: string) => void;
  onExport: (format: string) => void;
  onPrint: () => void;
  onCloseToHome: () => void;
  lastSavedTime?: string;
  wordCount?: number;
  cellCount?: number;
  docSize?: string;
}

export const FileMenuModal: React.FC<FileMenuModalProps> = ({
  isOpen,
  onClose,
  appMode,
  lang,
  plan,
  onTriggerGated,
  docTitle,
  onNew,
  onSave,
  onSaveAs,
  onExport,
  onPrint,
  onCloseToHome,
  lastSavedTime = 'Just now',
  wordCount = 0,
  cellCount = 0,
  docSize = '4.2 KB',
}) => {
  const [activeSubTab, setActiveSubTab] = useState<'info' | 'saveAs' | 'export'>('info');
  const [saveAsName, setSaveAsName] = useState(`${docTitle} - Copy`);
  const [saveToast, setSaveToast] = useState(false);
  const t = getTranslation(lang);

  if (!isOpen) return null;

  const isWord = appMode === 'word';
  const brandBg = isWord ? 'bg-[#185abd]' : 'bg-[#107c41]';
  const brandHover = isWord ? 'hover:bg-[#104a9e]' : 'hover:bg-[#0b5c30]';
  const brandText = isWord ? 'text-[#185abd]' : 'text-[#107c41]';

  const handleManualSave = () => {
    if (plan === 'free') {
      onTriggerGated('save_export');
      return;
    }
    onSave();
    setSaveToast(true);
    setTimeout(() => setSaveToast(false), 2000);
  };

  const handleSaveAsSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (plan === 'free') {
      onTriggerGated('save_export');
      return;
    }
    if (saveAsName.trim()) {
      onSaveAs(saveAsName.trim());
      onClose();
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex bg-white dark:bg-[#1f1e1d] text-gray-800 dark:text-gray-100 animate-in fade-in duration-150">
      {/* Left Backstage Rail */}
      <div className={`w-56 ${brandBg} text-white flex flex-col justify-between py-4 select-none shrink-0`}>
        <div className="space-y-1">
          {/* Back button */}
          <button
            onClick={onClose}
            className="flex items-center gap-3 px-5 py-3 hover:bg-black/20 w-full transition-colors text-sm font-medium"
          >
            <ArrowLeft className="w-4 h-4" />
            <span>{t.tabHome}</span>
          </button>

          <div className="pt-2 pb-1 px-5 text-xs font-semibold uppercase tracking-wider text-white/70">
            {t.fileMenuTitle}
          </div>

          <button
            onClick={() => {
              onNew();
              onClose();
            }}
            className="flex items-center gap-3 px-5 py-2.5 hover:bg-black/20 w-full transition-colors text-sm"
          >
            <FilePlus className="w-4 h-4" />
            <span>{t.newDocument}</span>
          </button>

          <button
            onClick={handleManualSave}
            className="flex items-center justify-between px-5 py-2.5 hover:bg-black/20 w-full transition-colors text-sm"
          >
            <span className="flex items-center gap-3">
              <Save className="w-4 h-4" />
              <span>{t.save}</span>
            </span>
            {plan === 'free' && <Lock className="w-3.5 h-3.5 text-amber-300" />}
          </button>

          <button
            onClick={() => {
              if (plan === 'free') {
                onTriggerGated('save_export');
              } else {
                setActiveSubTab('saveAs');
              }
            }}
            className={`flex items-center justify-between px-5 py-2.5 hover:bg-black/20 w-full transition-colors text-sm ${
              activeSubTab === 'saveAs' ? 'bg-black/25 font-semibold' : ''
            }`}
          >
            <span className="flex items-center gap-3">
              <Copy className="w-4 h-4" />
              <span>{t.saveAsDocument}</span>
            </span>
            {plan === 'free' && <Lock className="w-3.5 h-3.5 text-amber-300" />}
          </button>

          <button
            onClick={() => setActiveSubTab('export')}
            className={`flex items-center gap-3 px-5 py-2.5 hover:bg-black/20 w-full transition-colors text-sm ${
              activeSubTab === 'export' ? 'bg-black/25 font-semibold' : ''
            }`}
          >
            <Download className="w-4 h-4" />
            <span>{t.exportOptions}</span>
          </button>

          <button
            onClick={() => {
              onClose();
              onPrint();
            }}
            className="flex items-center gap-3 px-5 py-2.5 hover:bg-black/20 w-full transition-colors text-sm"
          >
            <Printer className="w-4 h-4" />
            <span>{t.print}</span>
          </button>

          <button
            onClick={() => setActiveSubTab('info')}
            className={`flex items-center gap-3 px-5 py-2.5 hover:bg-black/20 w-full transition-colors text-sm ${
              activeSubTab === 'info' ? 'bg-black/25 font-semibold' : ''
            }`}
          >
            <Info className="w-4 h-4" />
            <span>{t.documentProperties}</span>
          </button>
        </div>

        {/* Bottom launcher return */}
        <div className="px-3 pt-4 border-t border-white/20">
          <button
            onClick={onCloseToHome}
            className="flex items-center gap-3 px-3 py-2.5 hover:bg-black/20 w-full rounded text-sm text-red-100 hover:text-white transition-colors"
          >
            <LogOut className="w-4 h-4" />
            <span>{t.closeDoc}</span>
          </button>
        </div>
      </div>

      {/* Right Content Area */}
      <div className="flex-1 p-8 md:p-12 overflow-y-auto">
        {saveToast && (
          <div className="mb-6 p-3 rounded-md bg-green-500/10 border border-green-500/30 text-green-700 dark:text-green-300 flex items-center gap-2 text-sm">
            <CheckCircle2 className="w-4 h-4 text-green-500" />
            <span>{t.saved}</span>
          </div>
        )}

        {/* INFO TAB */}
        {activeSubTab === 'info' && (
          <div className="max-w-2xl space-y-6">
            <div>
              <h2 className="text-2xl font-semibold mb-1">{docTitle}</h2>
              <p className="text-sm text-gray-500 dark:text-gray-400">
                {isWord ? t.wordCardTitle : t.excelCardTitle}
              </p>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-3 gap-4 pt-4">
              <div className="p-4 rounded-lg bg-gray-50 dark:bg-[#292827] border border-gray-200 dark:border-[#3b3a39]">
                <div className="text-xs text-gray-500 dark:text-gray-400 mb-1">{t.fileSize}</div>
                <div className="text-lg font-semibold">{docSize}</div>
              </div>

              <div className="p-4 rounded-lg bg-gray-50 dark:bg-[#292827] border border-gray-200 dark:border-[#3b3a39]">
                <div className="text-xs text-gray-500 dark:text-gray-400 mb-1">
                  {isWord ? t.wordsCount : 'Cells Used'}
                </div>
                <div className="text-lg font-semibold">
                  {isWord ? wordCount.toLocaleString() : cellCount.toLocaleString()}
                </div>
              </div>

              <div className="p-4 rounded-lg bg-gray-50 dark:bg-[#292827] border border-gray-200 dark:border-[#3b3a39]">
                <div className="text-xs text-gray-500 dark:text-gray-400 mb-1">{t.lastModified}</div>
                <div className="text-sm font-semibold">{lastSavedTime}</div>
              </div>
            </div>

            <div className="p-5 rounded-lg border border-gray-200 dark:border-[#3b3a39] bg-white dark:bg-[#252423] space-y-3">
              <h3 className="font-semibold text-sm flex items-center gap-2">
                <Save className={`w-4 h-4 ${brandText}`} />
                <span>Local Persistence & Auto-Save</span>
              </h3>
              <p className="text-xs text-gray-600 dark:text-gray-300 leading-relaxed">
                Changes to this document are automatically saved locally in browser LocalStorage. You can also manually trigger save or export a standalone file anytime.
              </p>
              <div className="pt-2">
                <button
                  onClick={handleManualSave}
                  className={`px-4 py-2 rounded text-xs font-medium text-white ${brandBg} ${brandHover} transition-colors`}
                >
                  {t.save}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* SAVE AS TAB */}
        {activeSubTab === 'saveAs' && (
          <div className="max-w-md space-y-6">
            <h2 className="text-2xl font-semibold">{t.saveAsDocument}</h2>
            <form onSubmit={handleSaveAsSubmit} className="space-y-4">
              <div>
                <label className="block text-xs font-medium text-gray-600 dark:text-gray-300 mb-1">
                  New Document Title:
                </label>
                <input
                  type="text"
                  required
                  value={saveAsName}
                  onChange={(e) => setSaveAsName(e.target.value)}
                  className="w-full px-3 py-2 text-sm rounded border border-gray-300 dark:border-[#484644] bg-white dark:bg-[#252423] focus:outline-none focus:ring-2 focus:ring-[#185abd]"
                />
              </div>
              <button
                type="submit"
                className={`px-5 py-2 rounded text-xs font-medium text-white ${brandBg} ${brandHover} transition-colors`}
              >
                {t.saveAsDocument}
              </button>
            </form>
          </div>
        )}

        {/* EXPORT TAB */}
        {activeSubTab === 'export' && (
          <div className="max-w-2xl space-y-6">
            <div>
              <h2 className="text-2xl font-semibold mb-1">{t.exportOptions}</h2>
              <p className="text-sm text-gray-500 dark:text-gray-400">
                Choose a format to download and share your work.
              </p>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              {isWord ? (
                <>
                  <button
                    onClick={() => {
                      if (plan === 'free') {
                        onTriggerGated('save_export');
                      } else {
                        onExport('doc');
                      }
                    }}
                    className="p-4 text-left rounded-lg border border-gray-200 dark:border-[#3b3a39] hover:border-[#185abd] bg-white dark:bg-[#252423] hover:shadow-md transition-all group relative cursor-pointer"
                  >
                    <div className="flex items-center justify-between mb-2">
                      <div className="flex items-center gap-3">
                        <FileText className="w-6 h-6 text-[#185abd]" />
                        <span className="font-semibold text-sm group-hover:text-[#185abd]">Word Document (.docx / .doc)</span>
                      </div>
                      {plan === 'free' && <Lock className="w-4 h-4 text-amber-500" />}
                    </div>
                    <p className="text-xs text-gray-500 dark:text-gray-400">
                      Standard Microsoft Word compatible formatted rich document.
                    </p>
                  </button>

                  <button
                    onClick={() => {
                      if (plan === 'free') {
                        onTriggerGated('save_export');
                      } else {
                        onClose();
                        onPrint();
                      }
                    }}
                    className="p-4 text-left rounded-lg border border-gray-200 dark:border-[#3b3a39] hover:border-[#185abd] bg-white dark:bg-[#252423] hover:shadow-md transition-all group relative cursor-pointer"
                  >
                    <div className="flex items-center justify-between mb-2">
                      <div className="flex items-center gap-3">
                        <Printer className="w-6 h-6 text-[#185abd]" />
                        <span className="font-semibold text-sm group-hover:text-[#185abd]">PDF Document (.pdf)</span>
                      </div>
                      {plan === 'free' && <Lock className="w-4 h-4 text-amber-500" />}
                    </div>
                    <p className="text-xs text-gray-500 dark:text-gray-400">
                      Direct print preview formatted for crisp A4 paper output.
                    </p>
                  </button>

                  <button
                    onClick={() => onExport('html')}
                    className="p-4 text-left rounded-lg border border-gray-200 dark:border-[#3b3a39] hover:border-[#185abd] bg-white dark:bg-[#252423] hover:shadow-md transition-all group cursor-pointer"
                  >
                    <div className="flex items-center justify-between mb-2">
                      <div className="flex items-center gap-3">
                        <FileText className="w-6 h-6 text-[#185abd]" />
                        <span className="font-semibold text-sm group-hover:text-[#185abd]">Web Page (.html)</span>
                      </div>
                      <span className="text-[10px] uppercase font-bold text-emerald-600 bg-emerald-50 dark:bg-emerald-950/40 px-2 py-0.5 rounded">Free</span>
                    </div>
                    <p className="text-xs text-gray-500 dark:text-gray-400">
                      Clean semantic HTML suitable for web publication.
                    </p>
                  </button>

                  <button
                    onClick={() => onExport('txt')}
                    className="p-4 text-left rounded-lg border border-gray-200 dark:border-[#3b3a39] hover:border-[#185abd] bg-white dark:bg-[#252423] hover:shadow-md transition-all group cursor-pointer"
                  >
                    <div className="flex items-center justify-between mb-2">
                      <div className="flex items-center gap-3">
                        <FileText className="w-6 h-6 text-[#185abd]" />
                        <span className="font-semibold text-sm group-hover:text-[#185abd]">Plain Text (.txt)</span>
                      </div>
                      <span className="text-[10px] uppercase font-bold text-emerald-600 bg-emerald-50 dark:bg-emerald-950/40 px-2 py-0.5 rounded">Free</span>
                    </div>
                    <p className="text-xs text-gray-500 dark:text-gray-400">
                      Unformatted raw text document.
                    </p>
                  </button>
                </>
              ) : (
                <>
                  <button
                    onClick={() => {
                      if (plan === 'free') {
                        onTriggerGated('save_export');
                      } else {
                        onExport('csv');
                      }
                    }}
                    className="p-4 text-left rounded-lg border border-gray-200 dark:border-[#3b3a39] hover:border-[#107c41] bg-white dark:bg-[#252423] hover:shadow-md transition-all group relative cursor-pointer"
                  >
                    <div className="flex items-center justify-between mb-2">
                      <div className="flex items-center gap-3">
                        <FileSpreadsheet className="w-6 h-6 text-[#107c41]" />
                        <span className="font-semibold text-sm group-hover:text-[#107c41]">Excel Data (.xlsx / .csv)</span>
                      </div>
                      {plan === 'free' && <Lock className="w-4 h-4 text-amber-500" />}
                    </div>
                    <p className="text-xs text-gray-500 dark:text-gray-400">
                      Universal tabular spreadsheet data format.
                    </p>
                  </button>

                  <button
                    onClick={() => {
                      if (plan === 'free') {
                        onTriggerGated('save_export');
                      } else {
                        onExport('json');
                      }
                    }}
                    className="p-4 text-left rounded-lg border border-gray-200 dark:border-[#3b3a39] hover:border-[#107c41] bg-white dark:bg-[#252423] hover:shadow-md transition-all group relative cursor-pointer"
                  >
                    <div className="flex items-center justify-between mb-2">
                      <div className="flex items-center gap-3">
                        <FileSpreadsheet className="w-6 h-6 text-[#107c41]" />
                        <span className="font-semibold text-sm group-hover:text-[#107c41]">Workbook JSON (.json)</span>
                      </div>
                      {plan === 'free' && <Lock className="w-4 h-4 text-amber-500" />}
                    </div>
                    <p className="text-xs text-gray-500 dark:text-gray-400">
                      Full workbook representation including formulas, formatting, and sheets.
                    </p>
                  </button>

                  <button
                    onClick={() => onExport('html')}
                    className="p-4 text-left rounded-lg border border-gray-200 dark:border-[#3b3a39] hover:border-[#107c41] bg-white dark:bg-[#252423] hover:shadow-md transition-all group cursor-pointer"
                  >
                    <div className="flex items-center justify-between mb-2">
                      <div className="flex items-center gap-3">
                        <FileSpreadsheet className="w-6 h-6 text-[#107c41]" />
                        <span className="font-semibold text-sm group-hover:text-[#107c41]">HTML Table (.html)</span>
                      </div>
                      <span className="text-[10px] uppercase font-bold text-emerald-600 bg-emerald-50 dark:bg-emerald-950/40 px-2 py-0.5 rounded">Free</span>
                    </div>
                    <p className="text-xs text-gray-500 dark:text-gray-400">
                      Stylized HTML table ready for embedding.
                    </p>
                  </button>

                  <button
                    onClick={() => {
                      if (plan === 'free') {
                        onTriggerGated('save_export');
                      } else {
                        onClose();
                        onPrint();
                      }
                    }}
                    className="p-4 text-left rounded-lg border border-gray-200 dark:border-[#3b3a39] hover:border-[#107c41] bg-white dark:bg-[#252423] hover:shadow-md transition-all group relative cursor-pointer"
                  >
                    <div className="flex items-center justify-between mb-2">
                      <div className="flex items-center gap-3">
                        <Printer className="w-6 h-6 text-[#107c41]" />
                        <span className="font-semibold text-sm group-hover:text-[#107c41]">Print / Export PDF</span>
                      </div>
                      {plan === 'free' && <Lock className="w-4 h-4 text-amber-500" />}
                    </div>
                    <p className="text-xs text-gray-500 dark:text-gray-400">
                      Print clean grid view.
                    </p>
                  </button>
                </>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

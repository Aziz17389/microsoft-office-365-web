import React, { useRef, useState } from 'react';
import {
  FileText,
  FileSpreadsheet,
  Plus,
  Upload,
  Trash2,
  Clock,
  HardDrive,
  FileCode,
  ArrowRight,
  Globe,
  Sun,
  Moon,
  Sparkles,
  Crown,
  Lock,
  CheckCircle,
} from 'lucide-react';
import { AppMode, Language, PlanTier, RecentDocument, Theme } from '../../types/office';
import { getTranslation } from '../../i18n/translations';

interface OfficeLauncherProps {
  lang: Language;
  theme: Theme;
  plan: PlanTier;
  secondsRemaining: number;
  onOpenUpgradeModal: () => void;
  onTogglePlan: (plan: PlanTier) => void;
  recentDocs: RecentDocument[];
  onOpenWordTemplate: (templateKey: string) => void;
  onOpenExcelTemplate: (templateKey: string) => void;
  onOpenRecentDoc: (doc: RecentDocument) => void;
  onDeleteRecentDoc: (docId: string) => void;
  onImportFile: (file: File) => void;
  onToggleTheme: () => void;
  onToggleLanguage: () => void;
}

export const OfficeLauncher: React.FC<OfficeLauncherProps> = ({
  lang,
  theme,
  plan,
  secondsRemaining,
  onOpenUpgradeModal,
  onTogglePlan,
  recentDocs,
  onOpenWordTemplate,
  onOpenExcelTemplate,
  onOpenRecentDoc,
  onDeleteRecentDoc,
  onImportFile,
  onToggleTheme,
  onToggleLanguage,
}) => {
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const t = getTranslation(lang);

  const formatTime = (secs: number) => {
    const m = Math.floor(secs / 60);
    const s = secs % 60;
    return `${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
  };

  const wordTemplates = [
    { key: 'blank', title: t.templateBlankDoc, desc: 'Empty canvas' },
    { key: 'proposal', title: t.templateProposal, desc: 'Professional roadmap' },
    { key: 'meetingNotes', title: t.templateMeetingNotes, desc: 'Executive meeting' },
    { key: 'resume', title: t.templateResume, desc: 'Curriculum vitae' },
  ];

  const excelTemplates = [
    { key: 'blank', title: t.templateBlankSheet, desc: 'Empty grid' },
    { key: 'budget', title: t.templateBudget, desc: 'Income & Expenses' },
    { key: 'sales', title: t.templateSales, desc: 'Revenue analysis' },
    { key: 'gradebook', title: t.templateGradebook, desc: 'Grades & Pass/Fail' },
  ];

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      onImportFile(file);
    }
  };

  const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      onImportFile(file);
    }
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const formatDate = (timestamp: number) => {
    const date = new Date(timestamp);
    return date.toLocaleDateString(lang === 'ar' ? 'ar-SA' : 'en-US', {
      month: 'short',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });
  };

  return (
    <div className="min-h-screen bg-[#f3f2f1] dark:bg-[#11100f] text-gray-900 dark:text-gray-100 flex flex-col select-none transition-colors">
      {/* Top Navigation Bar */}
      <nav className="bg-[#185abd] dark:bg-[#1b1a19] text-white px-4 md:px-6 py-3 flex items-center justify-between shadow-md">
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-1 font-bold text-lg tracking-tight">
            <span className="w-6 h-6 rounded bg-white text-[#185abd] flex items-center justify-center font-black text-sm">
              ⊞
            </span>
            <span>{t.office365}</span>
          </div>
        </div>

        <div className="flex items-center gap-2">
          {/* Plan Badge & Timer */}
          {plan === 'free' ? (
            <button
              onClick={onOpenUpgradeModal}
              title="Click to upgrade to Office 365+"
              className="flex items-center gap-1.5 px-3 py-1 rounded-full bg-amber-400 text-black text-xs font-mono font-bold hover:bg-amber-300 transition-colors cursor-pointer"
            >
              <Clock className="w-3.5 h-3.5" />
              <span>{formatTime(secondsRemaining)}</span>
              <span className="hidden sm:inline font-sans text-[11px] font-normal">
                {t.timeRemaining}
              </span>
            </button>
          ) : (
            <button
              onClick={onOpenUpgradeModal}
              className="flex items-center gap-1.5 px-3 py-1 rounded-full bg-gradient-to-r from-amber-400 to-yellow-300 text-black text-xs font-black shadow-sm cursor-pointer hover:brightness-105 transition-all"
            >
              <Crown className="w-3.5 h-3.5" />
              <span>Office 365+ PRO</span>
            </button>
          )}

          {/* Test Switch Plan button */}
          <button
            onClick={() => onTogglePlan(plan === 'free' ? 'pro' : 'free')}
            title={t.demoToggle}
            className="hidden sm:flex items-center gap-1 px-2 py-1 rounded bg-white/10 hover:bg-white/20 text-xs font-mono text-white/90 transition-colors cursor-pointer"
          >
            <span>{plan === 'free' ? 'Free' : 'Pro'}</span>
            <span className="opacity-60">⇄</span>
          </button>

          {/* Language Switcher */}
          <button
            onClick={onToggleLanguage}
            title={lang === 'en' ? 'تبديل إلى العربية' : 'Switch to English'}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded hover:bg-white/15 text-xs font-semibold transition-colors cursor-pointer"
          >
            <Globe className="w-4 h-4" />
            <span>{lang === 'en' ? 'العربية' : 'English'}</span>
          </button>

          {/* Theme Switcher */}
          <button
            onClick={onToggleTheme}
            title={theme === 'dark' ? t.themeLight : t.themeDark}
            className="p-2 rounded hover:bg-white/15 transition-colors cursor-pointer"
          >
            {theme === 'dark' ? <Sun className="w-4 h-4 text-yellow-300" /> : <Moon className="w-4 h-4 text-white" />}
          </button>
        </div>
      </nav>

      {/* Main Content */}
      <main className="flex-1 max-w-6xl w-full mx-auto p-6 md:p-10 space-y-8">
        {/* Office 365+ Promotional Banner */}
        <div className="rounded-2xl p-6 bg-gradient-to-r from-amber-500/15 via-yellow-500/10 to-amber-500/15 border border-amber-300 dark:border-amber-500/30 shadow-sm flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
          <div className="space-y-1 max-w-2xl">
            <div className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full bg-amber-200 dark:bg-amber-950/60 text-amber-900 dark:text-amber-300 text-xs font-bold uppercase tracking-wider">
              <Crown className="w-3.5 h-3.5" />
              <span>Office 365+ Plan • $89.99 / month</span>
            </div>
            <h2 className="text-xl md:text-2xl font-black text-gray-900 dark:text-white">
              {t.bannerPromoTitle}
            </h2>
            <p className="text-xs md:text-sm text-gray-600 dark:text-gray-300 leading-relaxed">
              {t.bannerPromoDesc}
            </p>
          </div>

          <div className="flex items-center gap-3 shrink-0">
            {plan === 'free' ? (
              <button
                onClick={onOpenUpgradeModal}
                className="py-2.5 px-5 rounded-xl bg-gradient-to-r from-amber-500 to-amber-600 hover:from-amber-600 hover:to-amber-700 text-white font-bold text-xs md:text-sm shadow-md transition-all flex items-center gap-2 cursor-pointer"
              >
                <Crown className="w-4 h-4" />
                <span>{t.upgradeNow}</span>
              </button>
            ) : (
              <div className="flex items-center gap-2 px-4 py-2 rounded-xl bg-amber-400 text-black font-black text-xs md:text-sm shadow-sm">
                <CheckCircle className="w-4 h-4 text-black" />
                <span>Office 365+ Active</span>
              </div>
            )}
          </div>
        </div>

        {/* Welcome Header */}
        <div className="space-y-1">
          <h1 className="text-3xl md:text-4xl font-bold tracking-tight text-gray-900 dark:text-white">
            {t.welcomeTitle}
          </h1>
          <p className="text-sm md:text-base text-gray-600 dark:text-gray-400 max-w-3xl leading-relaxed">
            {t.welcomeSubtitle}
          </p>
        </div>

        {/* 2 Primary Application Cards */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {/* WORD CARD */}
          <div className="bg-white dark:bg-[#201f1e] rounded-xl p-6 shadow-sm hover:shadow-lg transition-all border border-gray-200 dark:border-[#3b3a39] flex flex-col justify-between group">
            <div className="space-y-4">
              <div className="flex items-start justify-between">
                <div className="flex items-center gap-3">
                  <div className="w-12 h-12 rounded-lg bg-[#185abd] text-white flex items-center justify-center font-bold text-2xl shadow-md">
                    W
                  </div>
                  <div>
                    <h2 className="text-xl font-bold text-gray-900 dark:text-white">
                      {t.wordCardTitle}
                    </h2>
                    <span className="text-xs text-[#185abd] dark:text-blue-400 font-medium">
                      Rich Text Document Processor
                    </span>
                  </div>
                </div>
              </div>

              <p className="text-xs md:text-sm text-gray-600 dark:text-gray-300 leading-relaxed">
                {t.wordCardDesc}
              </p>

              {/* Primary Action Button */}
              <button
                onClick={() => onOpenWordTemplate('blank')}
                className="w-full py-2.5 px-4 rounded-lg bg-[#185abd] hover:bg-[#104a9e] text-white font-medium text-sm flex items-center justify-center gap-2 shadow-sm transition-colors cursor-pointer"
              >
                <Plus className="w-4 h-4" />
                <span>{t.newBlankDocument}</span>
              </button>

              {/* Word Templates */}
              <div className="pt-3 border-t border-gray-100 dark:border-[#323130]">
                <div className="text-xs font-semibold text-gray-500 dark:text-gray-400 mb-2">
                  {t.templates}
                </div>
                <div className="grid grid-cols-2 gap-2">
                  {wordTemplates.slice(1).map((tpl) => (
                    <button
                      key={tpl.key}
                      onClick={() => onOpenWordTemplate(tpl.key)}
                      className="p-2.5 rounded-lg border border-gray-200 dark:border-[#3b3a39] hover:border-[#185abd] dark:hover:border-blue-500 hover:bg-blue-50/50 dark:hover:bg-blue-950/20 text-left rtl:text-right transition-colors cursor-pointer"
                    >
                      <div className="text-xs font-semibold text-gray-800 dark:text-gray-200 truncate">
                        {tpl.title}
                      </div>
                      <div className="text-[11px] text-gray-500 dark:text-gray-400 truncate">
                        {tpl.desc}
                      </div>
                    </button>
                  ))}
                </div>
              </div>
            </div>
          </div>

          {/* EXCEL CARD */}
          <div className="bg-white dark:bg-[#201f1e] rounded-xl p-6 shadow-sm hover:shadow-lg transition-all border border-gray-200 dark:border-[#3b3a39] flex flex-col justify-between group">
            <div className="space-y-4">
              <div className="flex items-start justify-between">
                <div className="flex items-center gap-3">
                  <div className="w-12 h-12 rounded-lg bg-[#107c41] text-white flex items-center justify-center font-bold text-2xl shadow-md">
                    X
                  </div>
                  <div>
                    <h2 className="text-xl font-bold text-gray-900 dark:text-white">
                      {t.excelCardTitle}
                    </h2>
                    <span className="text-xs text-[#107c41] dark:text-green-400 font-medium">
                      Intelligent Spreadsheet Engine
                    </span>
                  </div>
                </div>
              </div>

              <p className="text-xs md:text-sm text-gray-600 dark:text-gray-300 leading-relaxed">
                {t.excelCardDesc}
              </p>

              {/* Primary Action Button */}
              <button
                onClick={() => onOpenExcelTemplate('blank')}
                className="w-full py-2.5 px-4 rounded-lg bg-[#107c41] hover:bg-[#0b5c30] text-white font-medium text-sm flex items-center justify-center gap-2 shadow-sm transition-colors cursor-pointer"
              >
                <Plus className="w-4 h-4" />
                <span>{t.newBlankWorkbook}</span>
              </button>

              {/* Excel Templates */}
              <div className="pt-3 border-t border-gray-100 dark:border-[#323130]">
                <div className="text-xs font-semibold text-gray-500 dark:text-gray-400 mb-2">
                  {t.templates}
                </div>
                <div className="grid grid-cols-2 gap-2">
                  {excelTemplates.slice(1).map((tpl) => (
                    <button
                      key={tpl.key}
                      onClick={() => onOpenExcelTemplate(tpl.key)}
                      className="p-2.5 rounded-lg border border-gray-200 dark:border-[#3b3a39] hover:border-[#107c41] dark:hover:border-green-500 hover:bg-green-50/50 dark:hover:bg-green-950/20 text-left rtl:text-right transition-colors cursor-pointer"
                    >
                      <div className="text-xs font-semibold text-gray-800 dark:text-gray-200 truncate">
                        {tpl.title}
                      </div>
                      <div className="text-[11px] text-gray-500 dark:text-gray-400 truncate">
                        {tpl.desc}
                      </div>
                    </button>
                  ))}
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Drag-and-Drop / File Import Section */}
        <div
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
          className={`p-6 rounded-xl border-2 border-dashed transition-all cursor-pointer flex flex-col items-center justify-center text-center ${
            isDragging
              ? 'border-[#185abd] bg-blue-50/50 dark:bg-blue-950/20 scale-[1.01]'
              : 'border-gray-300 dark:border-[#3b3a39] bg-white dark:bg-[#201f1e] hover:border-gray-400 dark:hover:border-[#484644]'
          }`}
        >
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileInputChange}
            accept=".docx,.xlsx,.csv,.txt,.html"
            className="hidden"
          />
          <Upload className="w-8 h-8 text-[#185abd] dark:text-blue-400 mb-2" />
          <h3 className="font-semibold text-sm mb-1">{t.importFile}</h3>
          <p className="text-xs text-gray-500 dark:text-gray-400 max-w-sm">
            {t.importDesc}
          </p>
        </div>

        {/* Recent Documents Section */}
        <div className="space-y-4">
          <div className="flex items-center justify-between">
            <h2 className="text-xl font-bold flex items-center gap-2">
              <Clock className="w-5 h-5 text-gray-500" />
              <span>{t.recentFiles}</span>
            </h2>
            <div className="flex items-center gap-2">
              <span className="text-xs text-gray-500 font-mono">
                {recentDocs.length} saved
              </span>
            </div>
          </div>

          {recentDocs.length === 0 ? (
            <div className="p-8 text-center bg-white dark:bg-[#201f1e] rounded-xl border border-gray-200 dark:border-[#3b3a39] text-gray-500 text-xs">
              {t.noRecentFiles}
            </div>
          ) : (
            <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
              {recentDocs.map((doc) => {
                const isWord = doc.app === 'word';
                return (
                  <div
                    key={doc.id}
                    onClick={() => onOpenRecentDoc(doc)}
                    className="p-4 bg-white dark:bg-[#201f1e] rounded-lg border border-gray-200 dark:border-[#3b3a39] hover:border-gray-400 dark:hover:border-gray-500 shadow-xs hover:shadow-md transition-all flex items-center justify-between group cursor-pointer"
                  >
                    <div className="flex items-center gap-3 overflow-hidden">
                      <div
                        className={`w-9 h-9 rounded-md flex items-center justify-center font-bold text-white shrink-0 ${
                          isWord ? 'bg-[#185abd]' : 'bg-[#107c41]'
                        }`}
                      >
                        {isWord ? 'W' : 'X'}
                      </div>
                      <div className="overflow-hidden">
                        <div className="font-semibold text-sm text-gray-900 dark:text-gray-100 truncate group-hover:text-[#185abd] dark:group-hover:text-blue-400 transition-colors">
                          {doc.title}
                        </div>
                        <div className="text-[11px] text-gray-500 dark:text-gray-400 truncate mt-0.5">
                          {doc.preview}
                        </div>
                        <div className="flex items-center gap-3 text-[10px] text-gray-400 mt-1 font-mono">
                          <span>{formatDate(doc.updatedAt)}</span>
                          <span>•</span>
                          <span>{doc.size}</span>
                        </div>
                      </div>
                    </div>

                    {/* Actions */}
                    <div className="flex items-center gap-1 shrink-0 ml-2 rtl:ml-0 rtl:mr-2">
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          if (window.confirm(t.confirmDelete)) {
                            onDeleteRecentDoc(doc.id);
                          }
                        }}
                        title={t.deleteFile}
                        className="p-1.5 rounded hover:bg-red-50 dark:hover:bg-red-950/40 text-gray-400 hover:text-red-500 transition-colors"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </main>
    </div>
  );
};

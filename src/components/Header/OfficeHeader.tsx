import React, { useState, useRef, useEffect } from 'react';
import {
  Grid,
  Save,
  Undo2,
  Redo2,
  Search,
  Sun,
  Moon,
  Globe,
  Maximize2,
  Home,
  Check,
  Crown,
  Clock,
  Lock,
  Sparkles,
} from 'lucide-react';
import { AppMode, Language, PlanTier, Theme } from '../../types/office';
import { getTranslation } from '../../i18n/translations';

interface OfficeHeaderProps {
  appMode: AppMode;
  lang: Language;
  theme: Theme;
  docTitle: string;
  isSaved: boolean;
  canUndo?: boolean;
  canRedo?: boolean;
  plan: PlanTier;
  secondsRemaining: number;
  onOpenUpgradeModal: () => void;
  onTogglePlan: (plan: PlanTier) => void;
  onTitleChange: (newTitle: string) => void;
  onSave: () => void;
  onUndo: () => void;
  onRedo: () => void;
  onSwitchApp: (mode: AppMode) => void;
  onToggleTheme: () => void;
  onToggleLanguage: () => void;
  onToggleZenMode: () => void;
  onExecuteCommand?: (cmd: string) => void;
}

export const OfficeHeader: React.FC<OfficeHeaderProps> = ({
  appMode,
  lang,
  theme,
  docTitle,
  isSaved,
  canUndo = true,
  canRedo = true,
  plan,
  secondsRemaining,
  onOpenUpgradeModal,
  onTogglePlan,
  onTitleChange,
  onSave,
  onUndo,
  onRedo,
  onSwitchApp,
  onToggleTheme,
  onToggleLanguage,
  onToggleZenMode,
  onExecuteCommand,
}) => {
  const [isWaffleOpen, setIsWaffleOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [isSearchFocused, setIsSearchFocused] = useState(false);
  const waffleRef = useRef<HTMLDivElement>(null);
  const searchRef = useRef<HTMLDivElement>(null);

  const t = getTranslation(lang);
  const isWord = appMode === 'word';
  const isExcel = appMode === 'excel';

  // Format seconds to MM:SS
  const formatTime = (secs: number) => {
    const m = Math.floor(secs / 60);
    const s = secs % 60;
    return `${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
  };

  // Office branding colors
  const brandBg = isWord
    ? 'bg-[#185abd] dark:bg-[#0f3d82]'
    : isExcel
    ? 'bg-[#107c41] dark:bg-[#0b542c]'
    : 'bg-[#2b579a] dark:bg-[#1f1e1d]';

  // Close menus on outside click
  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (waffleRef.current && !waffleRef.current.contains(event.target as Node)) {
        setIsWaffleOpen(false);
      }
      if (searchRef.current && !searchRef.current.contains(event.target as Node)) {
        setIsSearchFocused(false);
      }
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Quick command suggestions for "Tell me what you want to do"
  const searchActions = [
    { id: 'bold', label: lang === 'ar' ? 'غامق (Bold)' : 'Bold text', icon: 'B', app: 'word' },
    { id: 'table', label: lang === 'ar' ? 'إدراج جدول (Table) 🔒' : 'Insert Table 🔒', icon: '⊞', app: 'any' },
    { id: 'find', label: lang === 'ar' ? 'بحث واستبدال (Find & Replace)' : 'Find & Replace', icon: '🔍', app: 'any' },
    { id: 'print', label: lang === 'ar' ? 'طباعة المستند (Print) 🔒' : 'Print Document 🔒', icon: '🖨️', app: 'any' },
    { id: 'autosum', label: lang === 'ar' ? 'جمع تلقائي (AutoSum)' : 'AutoSum Formula', icon: '∑', app: 'excel' },
    { id: 'currency', label: lang === 'ar' ? 'تنسيق العملة (Currency)' : 'Format as Currency', icon: '$', app: 'excel' },
    { id: 'upgrade', label: lang === 'ar' ? 'الترقية إلى Office 365+' : 'Upgrade to Office 365+', icon: '👑', app: 'any' },
    { id: 'darkmode', label: lang === 'ar' ? 'تبديل المظهر الداكن' : 'Toggle Dark Mode', icon: '🌓', app: 'any' },
  ].filter((a) => a.app === 'any' || a.app === appMode);

  const filteredActions = searchQuery.trim()
    ? searchActions.filter((a) => a.label.toLowerCase().includes(searchQuery.toLowerCase()))
    : searchActions;

  return (
    <header className={`${brandBg} text-white px-2 py-1.5 flex items-center justify-between shadow-md select-none relative z-40 transition-colors`}>
      {/* Left section: Waffle launcher, title, Quick Access Toolbar */}
      <div className="flex items-center gap-2">
        {/* 9-dot Waffle Menu */}
        <div className="relative" ref={waffleRef}>
          <button
            onClick={() => setIsWaffleOpen(!isWaffleOpen)}
            title={t.waffleMenu}
            className="p-1.5 rounded hover:bg-white/15 active:bg-white/25 transition-colors flex items-center justify-center cursor-pointer"
          >
            <Grid className="w-5 h-5 text-white" />
          </button>

          {/* Quick switcher popover */}
          {isWaffleOpen && (
            <div className="absolute top-full mt-1.5 left-0 rtl:left-auto rtl:right-0 w-64 bg-white dark:bg-[#252423] text-gray-800 dark:text-gray-100 rounded-lg shadow-2xl border border-gray-200 dark:border-[#3b3a39] p-3 z-50 animate-in fade-in duration-100">
              <div className="text-xs font-semibold text-gray-500 dark:text-gray-400 mb-2 px-1">
                {t.switchAppPrompt}
              </div>
              <div className="space-y-1">
                <button
                  onClick={() => {
                    onSwitchApp('word');
                    setIsWaffleOpen(false);
                  }}
                  className={`w-full flex items-center gap-3 p-2 rounded-md hover:bg-gray-100 dark:hover:bg-[#323130] transition-colors text-left rtl:text-right ${
                    isWord ? 'bg-blue-50 dark:bg-blue-950/40 text-[#185abd] dark:text-blue-400 font-medium' : ''
                  }`}
                >
                  <div className="w-8 h-8 rounded bg-[#185abd] text-white flex items-center justify-center font-bold text-sm shrink-0">
                    W
                  </div>
                  <div>
                    <div className="text-xs font-semibold">{t.wordApp}</div>
                    <div className="text-[11px] text-gray-500 dark:text-gray-400">Document Editor</div>
                  </div>
                </button>

                <button
                  onClick={() => {
                    onSwitchApp('excel');
                    setIsWaffleOpen(false);
                  }}
                  className={`w-full flex items-center gap-3 p-2 rounded-md hover:bg-gray-100 dark:hover:bg-[#323130] transition-colors text-left rtl:text-right ${
                    isExcel ? 'bg-green-50 dark:bg-green-950/40 text-[#107c41] dark:text-green-400 font-medium' : ''
                  }`}
                >
                  <div className="w-8 h-8 rounded bg-[#107c41] text-white flex items-center justify-center font-bold text-sm shrink-0">
                    X
                  </div>
                  <div>
                    <div className="text-xs font-semibold">{t.excelApp}</div>
                    <div className="text-[11px] text-gray-500 dark:text-gray-400">Spreadsheet Engine</div>
                  </div>
                </button>

                <div className="pt-2 border-t border-gray-200 dark:border-[#3b3a39]">
                  <button
                    onClick={() => {
                      onSwitchApp('launcher');
                      setIsWaffleOpen(false);
                    }}
                    className="w-full flex items-center gap-3 p-2 rounded-md hover:bg-gray-100 dark:hover:bg-[#323130] transition-colors text-left rtl:text-right text-gray-700 dark:text-gray-300"
                  >
                    <Home className="w-4 h-4 text-gray-500" />
                    <span className="text-xs">{t.backToHome}</span>
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Office Branding Icon */}
        <div className="flex items-center gap-2 pr-1 rtl:pr-0 rtl:pl-1">
          {isWord ? (
            <div className="w-6 h-6 rounded bg-white/20 flex items-center justify-center font-bold text-xs tracking-tighter">
              W
            </div>
          ) : isExcel ? (
            <div className="w-6 h-6 rounded bg-white/20 flex items-center justify-center font-bold text-xs tracking-tighter">
              X
            </div>
          ) : null}
          <span className="font-semibold text-sm hidden sm:inline-block tracking-tight">
            {isWord ? 'Word' : isExcel ? 'Excel' : 'Office 365'}
          </span>
        </div>

        {/* Document Title (Inline click-to-edit) */}
        <div className="flex items-center gap-1.5 pl-1 rtl:pl-0 rtl:pr-1 max-w-[170px] sm:max-w-xs">
          <input
            type="text"
            value={docTitle}
            onChange={(e) => onTitleChange(e.target.value)}
            className="bg-transparent hover:bg-white/10 focus:bg-white/20 text-white font-medium text-xs sm:text-sm px-2 py-0.5 rounded outline-none border border-transparent focus:border-white/30 truncate transition-colors"
            title="Click to rename document"
          />
        </div>

        {/* Quick Access Toolbar */}
        <div className="flex items-center gap-0.5 border-l rtl:border-l-0 rtl:border-r border-white/20 pl-2 rtl:pl-0 rtl:pr-2">
          {/* Manual Save (Gated behind Office 365+) */}
          <button
            onClick={plan === 'pro' ? onSave : onOpenUpgradeModal}
            title={plan === 'pro' ? t.save : `${t.save} (${t.lockedFeatureTooltip})`}
            className="p-1 rounded hover:bg-white/15 transition-colors relative group"
          >
            <Save className="w-4 h-4 text-white" />
            {plan === 'free' && (
              <span className="absolute -top-1 -right-1 rtl:-right-auto rtl:-left-1 bg-amber-400 text-black text-[9px] rounded-full p-0.5 font-bold shadow-xs">
                <Lock className="w-2.5 h-2.5" />
              </span>
            )}
          </button>

          {/* Undo */}
          <button
            onClick={onUndo}
            disabled={!canUndo}
            title={t.undo}
            className="p-1 rounded hover:bg-white/15 disabled:opacity-40 transition-colors"
          >
            <Undo2 className="w-4 h-4 text-white" />
          </button>

          {/* Redo */}
          <button
            onClick={onRedo}
            disabled={!canRedo}
            title={t.redo}
            className="p-1 rounded hover:bg-white/15 disabled:opacity-40 transition-colors"
          >
            <Redo2 className="w-4 h-4 text-white" />
          </button>

          {/* Save Status Pill (Pro only) */}
          {plan === 'pro' && (
            <div className="hidden xl:flex items-center gap-1 text-[11px] text-white/80 pl-2 rtl:pl-0 rtl:pr-2">
              {isSaved ? (
                <span className="flex items-center gap-1">
                  <Check className="w-3 h-3 text-white" />
                  <span className="opacity-80">{t.saved}</span>
                </span>
              ) : (
                <span className="opacity-80">{t.saving}</span>
              )}
            </div>
          )}
        </div>
      </div>

      {/* Center Search Bar: "Tell me what you want to do..." */}
      <div className="hidden lg:flex flex-1 max-w-xs xl:max-w-md mx-3 relative" ref={searchRef}>
        <div className="relative w-full">
          <Search className="w-3.5 h-3.5 absolute left-2.5 rtl:left-auto rtl:right-2.5 top-1/2 -translate-y-1/2 text-white/70 pointer-events-none" />
          <input
            type="text"
            value={searchQuery}
            onFocus={() => setIsSearchFocused(true)}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder={t.searchPlaceholder}
            className="w-full bg-white/15 hover:bg-white/20 focus:bg-white focus:text-gray-900 placeholder:text-white/70 focus:placeholder:text-gray-400 pl-8 rtl:pl-3 rtl:pr-8 pr-3 py-1 rounded text-xs transition-all outline-none"
          />
        </div>

        {/* Search quick actions popover */}
        {isSearchFocused && (
          <div className="absolute top-full mt-1.5 left-0 right-0 bg-white dark:bg-[#252423] text-gray-800 dark:text-gray-100 rounded-lg shadow-xl border border-gray-200 dark:border-[#3b3a39] p-2 z-50">
            <div className="text-[10px] font-semibold text-gray-400 uppercase tracking-wider px-2 py-1">
              Suggested Actions
            </div>
            <div className="space-y-0.5">
              {filteredActions.map((action) => (
                <button
                  key={action.id}
                  onClick={() => {
                    if (action.id === 'upgrade') onOpenUpgradeModal();
                    else if (onExecuteCommand) onExecuteCommand(action.id);
                    if (action.id === 'darkmode') onToggleTheme();
                    setIsSearchFocused(false);
                    setSearchQuery('');
                  }}
                  className="w-full text-left rtl:text-right px-2.5 py-1.5 rounded hover:bg-gray-100 dark:hover:bg-[#323130] flex items-center justify-between text-xs transition-colors"
                >
                  <span className="flex items-center gap-2">
                    <span className="w-5 h-5 rounded bg-gray-100 dark:bg-[#323130] flex items-center justify-center font-bold text-[10px]">
                      {action.icon}
                    </span>
                    <span>{action.label}</span>
                  </span>
                  <span className="text-[10px] text-gray-400">Action</span>
                </button>
              ))}
            </div>
          </div>
        )}
      </div>

      {/* Right section: Subscription Status Badge / Timer, Utilities */}
      <div className="flex items-center gap-1.5">
        {/* Plan Status & Live 10-Min Countdown Timer */}
        {plan === 'free' ? (
          <div className="flex items-center gap-1">
            {/* Live Floating Timer Button */}
            <button
              onClick={onOpenUpgradeModal}
              title="10-Minute Session Limit. Click to upgrade."
              className={`flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-mono font-bold shadow-xs transition-all cursor-pointer ${
                secondsRemaining <= 120
                  ? 'bg-rose-500 text-white animate-pulse'
                  : 'bg-amber-400/95 text-black hover:bg-amber-300'
              }`}
            >
              <Clock className="w-3.5 h-3.5 shrink-0" />
              <span>{formatTime(secondsRemaining)}</span>
              <span className="hidden sm:inline font-sans text-[11px] font-normal opacity-90">
                {t.timeRemaining}
              </span>
            </button>

            {/* Quick Upgrade Button */}
            <button
              onClick={onOpenUpgradeModal}
              className="hidden sm:flex items-center gap-1 px-2.5 py-1 rounded bg-amber-500 hover:bg-amber-400 text-white text-xs font-bold transition-all shadow-xs cursor-pointer"
            >
              <Crown className="w-3.5 h-3.5" />
              <span>$89.99/mo</span>
            </button>
          </div>
        ) : (
          /* Pro Plan Badge */
          <div className="flex items-center gap-1">
            <button
              onClick={onOpenUpgradeModal}
              title="Office 365+ Active with Unlimited Time & Features"
              className="flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-gradient-to-r from-amber-400 to-yellow-300 text-black text-xs font-black shadow-md cursor-pointer hover:brightness-105 transition-all"
            >
              <Crown className="w-3.5 h-3.5 text-black" />
              <span>Office 365+ PRO</span>
            </button>
          </div>
        )}

        {/* Demo Switch Plan button */}
        <button
          onClick={() => onTogglePlan(plan === 'free' ? 'pro' : 'free')}
          title={t.demoToggle}
          className="hidden md:flex items-center gap-1 px-1.5 py-0.5 rounded bg-white/10 hover:bg-white/20 text-[10px] text-white/90 font-mono transition-colors cursor-pointer"
        >
          <span>{plan === 'free' ? 'Free' : 'Pro'}</span>
          <span className="opacity-60">⇄</span>
        </button>

        {/* Language Switcher */}
        <button
          onClick={onToggleLanguage}
          title={lang === 'en' ? 'تبديل إلى العربية' : 'Switch to English'}
          className="flex items-center gap-1 px-2 py-1 rounded hover:bg-white/15 text-xs font-medium transition-colors cursor-pointer"
        >
          <Globe className="w-3.5 h-3.5" />
          <span className="uppercase">{lang === 'en' ? 'عربي' : 'EN'}</span>
        </button>

        {/* Theme Switcher */}
        <button
          onClick={onToggleTheme}
          title={theme === 'dark' ? t.themeLight : t.themeDark}
          className="p-1.5 rounded hover:bg-white/15 transition-colors cursor-pointer"
        >
          {theme === 'dark' ? <Sun className="w-4 h-4 text-yellow-300" /> : <Moon className="w-4 h-4 text-white" />}
        </button>

        {/* Fullscreen / Zen Mode */}
        <button
          onClick={onToggleZenMode}
          title={t.fullscreen}
          className="p-1.5 rounded hover:bg-white/15 transition-colors cursor-pointer"
        >
          <Maximize2 className="w-4 h-4" />
        </button>

        {/* Home / Launcher return */}
        <button
          onClick={() => onSwitchApp('launcher')}
          title={t.backToHome}
          className="p-1.5 rounded hover:bg-white/15 transition-colors cursor-pointer"
        >
          <Home className="w-4 h-4" />
        </button>
      </div>
    </header>
  );
};


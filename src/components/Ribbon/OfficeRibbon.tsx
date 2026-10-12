import React from 'react';
import { AppMode, Language, RibbonTab } from '../../types/office';
import { getTranslation } from '../../i18n/translations';

interface OfficeRibbonProps {
  appMode: AppMode;
  lang: Language;
  activeTab: RibbonTab;
  onSelectTab: (tab: RibbonTab) => void;
  onOpenFileMenu: () => void;
  children: React.ReactNode;
}

export const OfficeRibbon: React.FC<OfficeRibbonProps> = ({
  appMode,
  lang,
  activeTab,
  onSelectTab,
  onOpenFileMenu,
  children,
}) => {
  const t = getTranslation(lang);
  const isWord = appMode === 'word';
  const fileTabBg = isWord ? 'bg-[#185abd]' : 'bg-[#107c41]';
  const activeTabBorder = isWord
    ? 'border-[#185abd] text-[#185abd] dark:text-blue-400 dark:border-blue-400'
    : 'border-[#107c41] text-[#107c41] dark:text-green-400 dark:border-green-400';

  const tabs: { id: RibbonTab; label: string }[] = [
    { id: 'home', label: t.tabHome },
    { id: 'insert', label: t.tabInsert },
    { id: 'layout', label: t.tabLayout },
    { id: 'view', label: t.tabView },
  ];

  return (
    <div className="ribbon-container bg-[#f3f2f1] dark:bg-[#201f1e] border-b border-gray-300 dark:border-[#3b3a39] shadow-sm select-none">
      {/* Ribbon Tab Bar Header */}
      <div className="flex items-center px-2 pt-1 border-b border-gray-200 dark:border-[#2d2c2b]">
        {/* File Tab (Special Backstage button) */}
        <button
          onClick={onOpenFileMenu}
          className={`${fileTabBg} text-white font-medium text-xs px-3.5 py-1.5 rounded-t hover:brightness-110 active:brightness-95 transition-all mr-1 cursor-pointer`}
        >
          {t.tabFile}
        </button>

        {/* Standard Ribbon Tabs */}
        <div className="flex items-center gap-0.5">
          {tabs.map((tab) => {
            const isActive = activeTab === tab.id;
            return (
              <button
                key={tab.id}
                onClick={() => onSelectTab(tab.id)}
                className={`px-3 py-1.5 text-xs font-medium border-b-2 transition-colors cursor-pointer ${
                  isActive
                    ? `${activeTabBorder} bg-white dark:bg-[#252423] rounded-t font-semibold`
                    : 'border-transparent text-gray-700 dark:text-gray-300 hover:bg-gray-200 dark:hover:bg-[#2d2c2b]'
                }`}
              >
                {tab.label}
              </button>
            );
          })}
        </div>
      </div>

      {/* Ribbon Tab Body Container */}
      <div className="bg-white dark:bg-[#252423] py-1 border-b border-gray-300 dark:border-[#3b3a39]">
        {children}
      </div>
    </div>
  );
};

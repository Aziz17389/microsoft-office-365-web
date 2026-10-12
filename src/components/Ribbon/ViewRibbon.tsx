import React from 'react';
import {
  Ruler,
  Grid,
  Maximize2,
  ZoomIn,
} from 'lucide-react';
import { Language } from '../../types/office';
import { getTranslation } from '../../i18n/translations';

interface ViewRibbonProps {
  lang: Language;
  showRuler: boolean;
  showGridlines: boolean;
  isZenMode: boolean;
  zoom: number;
  onToggleRuler: () => void;
  onToggleGridlines: () => void;
  onToggleZenMode: () => void;
  onSetZoom: (zoom: number) => void;
}

export const ViewRibbon: React.FC<ViewRibbonProps> = ({
  lang,
  showRuler,
  showGridlines,
  isZenMode,
  zoom,
  onToggleRuler,
  onToggleGridlines,
  onToggleZenMode,
  onSetZoom,
}) => {
  const t = getTranslation(lang);
  const zoomLevels = [50, 75, 100, 125, 150, 200];

  return (
    <div className="flex items-stretch h-24 overflow-x-auto select-none px-2 py-1 text-gray-700 dark:text-gray-200">
      {/* 1. SHOW / HIDE GROUP */}
      <div className="flex flex-col justify-between px-3 border-r border-gray-200 dark:border-[#3b3a39] shrink-0">
        <div className="space-y-1.5 my-auto">
          <label className="flex items-center gap-2 text-xs cursor-pointer select-none">
            <input
              type="checkbox"
              checked={showRuler}
              onChange={onToggleRuler}
              className="rounded border-gray-300 text-[#185abd] focus:ring-[#185abd]"
            />
            <span className="flex items-center gap-1.5">
              <Ruler className="w-3.5 h-3.5 text-gray-500" />
              <span>{t.ruler}</span>
            </span>
          </label>

          <label className="flex items-center gap-2 text-xs cursor-pointer select-none">
            <input
              type="checkbox"
              checked={showGridlines}
              onChange={onToggleGridlines}
              className="rounded border-gray-300 text-[#185abd] focus:ring-[#185abd]"
            />
            <span className="flex items-center gap-1.5">
              <Grid className="w-3.5 h-3.5 text-gray-500" />
              <span>{t.gridlines}</span>
            </span>
          </label>
        </div>
        <div className="text-[10px] text-center text-gray-400 dark:text-gray-500 font-medium">
          {t.groupShow}
        </div>
      </div>

      {/* 2. ZEN MODE / FULLSCREEN */}
      <div className="flex flex-col justify-between px-3 border-r border-gray-200 dark:border-[#3b3a39] shrink-0">
        <div className="flex items-center my-auto">
          <button
            onClick={onToggleZenMode}
            className={`flex flex-col items-center justify-center p-2 rounded transition-colors w-16 ${
              isZenMode ? 'bg-[#185abd] text-white' : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
            }`}
          >
            <Maximize2 className="w-5 h-5 mb-1" />
            <span className="text-[11px] font-medium">{t.fullscreen}</span>
          </button>
        </div>
        <div className="text-[10px] text-center text-gray-400 dark:text-gray-500 font-medium">
          Mode
        </div>
      </div>

      {/* 3. ZOOM PRESETS */}
      <div className="flex flex-col justify-between px-3 shrink-0">
        <div className="flex items-center gap-2 my-auto">
          <div className="flex flex-col items-center">
            <ZoomIn className="w-5 h-5 text-gray-600 dark:text-gray-300 mb-1" />
            <span className="text-[10px] text-gray-500 font-mono">{zoom}%</span>
          </div>

          <div className="grid grid-cols-3 gap-1">
            {zoomLevels.map((lvl) => (
              <button
                key={lvl}
                onClick={() => onSetZoom(lvl)}
                className={`px-2 py-0.5 text-xs rounded font-mono transition-colors ${
                  zoom === lvl
                    ? 'bg-[#185abd] text-white font-semibold'
                    : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
                }`}
              >
                {lvl}%
              </button>
            ))}
          </div>
        </div>
        <div className="text-[10px] text-center text-gray-400 dark:text-gray-500 font-medium">
          {t.groupZoom}
        </div>
      </div>
    </div>
  );
};

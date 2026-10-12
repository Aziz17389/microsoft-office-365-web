import React, { useState } from 'react';
import { X, Table as TableIcon } from 'lucide-react';
import { Language } from '../../types/office';
import { getTranslation } from '../../i18n/translations';

interface InsertTableModalProps {
  isOpen: boolean;
  onClose: () => void;
  lang: Language;
  onInsertTable: (rows: number, cols: number) => void;
}

export const InsertTableModal: React.FC<InsertTableModalProps> = ({
  isOpen,
  onClose,
  lang,
  onInsertTable,
}) => {
  const [hoveredRows, setHoveredRows] = useState(3);
  const [hoveredCols, setHoveredCols] = useState(3);
  const [customRows, setCustomRows] = useState(3);
  const [customCols, setCustomCols] = useState(3);
  const t = getTranslation(lang);

  if (!isOpen) return null;

  const MAX_GRID = 8;

  const handleGridHover = (r: number, c: number) => {
    setHoveredRows(r);
    setHoveredCols(c);
    setCustomRows(r);
    setCustomCols(c);
  };

  const handleConfirm = () => {
    const r = Math.max(1, Math.min(50, customRows));
    const c = Math.max(1, Math.min(20, customCols));
    onInsertTable(r, c);
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/45 backdrop-blur-[1px] p-4">
      <div className="w-full max-w-sm rounded-lg bg-white dark:bg-[#252423] text-gray-800 dark:text-gray-100 shadow-2xl border border-gray-300 dark:border-[#3b3a39] overflow-hidden">
        {/* Header */}
        <div className="flex items-center justify-between px-4 py-3 bg-gray-100 dark:bg-[#1f1e1d] border-b border-gray-200 dark:border-[#3b3a39]">
          <div className="flex items-center gap-2">
            <TableIcon className="w-4 h-4 text-[#185abd] dark:text-blue-400" />
            <h3 className="font-semibold text-sm">{t.tableDimensions}</h3>
          </div>
          <button onClick={onClose} className="p-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130]">
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Content */}
        <div className="p-4 space-y-4">
          {/* Interactive Visual Grid */}
          <div className="flex flex-col items-center">
            <div className="text-xs font-semibold text-[#185abd] dark:text-blue-400 mb-2">
              {hoveredCols} × {hoveredRows} {t.insertTable}
            </div>
            <div
              className="grid gap-1 p-2 bg-gray-50 dark:bg-[#1b1a19] rounded border border-gray-200 dark:border-[#3b3a39]"
              style={{ gridTemplateColumns: `repeat(${MAX_GRID}, minmax(0, 1fr))` }}
            >
              {Array.from({ length: MAX_GRID }).map((_, rIdx) => {
                const r = rIdx + 1;
                return Array.from({ length: MAX_GRID }).map((_, cIdx) => {
                  const c = cIdx + 1;
                  const isSelected = r <= hoveredRows && c <= hoveredCols;
                  return (
                    <div
                      key={`${r}-${c}`}
                      onMouseEnter={() => handleGridHover(r, c)}
                      onClick={handleConfirm}
                      className={`w-5 h-5 rounded-[2px] border cursor-pointer transition-colors ${
                        isSelected
                          ? 'bg-[#185abd]/30 border-[#185abd]'
                          : 'bg-white dark:bg-[#252423] border-gray-300 dark:border-gray-700 hover:border-gray-400'
                      }`}
                    />
                  );
                });
              })}
            </div>
          </div>

          {/* Numeric dimensions */}
          <div className="grid grid-cols-2 gap-3 pt-2 border-t border-gray-200 dark:border-[#3b3a39]">
            <div>
              <label className="block text-xs text-gray-500 dark:text-gray-400 mb-1">
                {t.tableCols}
              </label>
              <input
                type="number"
                min="1"
                max="20"
                value={customCols}
                onChange={(e) => {
                  const v = parseInt(e.target.value, 10) || 1;
                  setCustomCols(v);
                  if (v <= MAX_GRID) setHoveredCols(v);
                }}
                className="w-full px-2 py-1 text-sm rounded border border-gray-300 dark:border-[#484644] bg-white dark:bg-[#1b1a19]"
              />
            </div>
            <div>
              <label className="block text-xs text-gray-500 dark:text-gray-400 mb-1">
                {t.tableRows}
              </label>
              <input
                type="number"
                min="1"
                max="50"
                value={customRows}
                onChange={(e) => {
                  const v = parseInt(e.target.value, 10) || 1;
                  setCustomRows(v);
                  if (v <= MAX_GRID) setHoveredRows(v);
                }}
                className="w-full px-2 py-1 text-sm rounded border border-gray-300 dark:border-[#484644] bg-white dark:bg-[#1b1a19]"
              />
            </div>
          </div>
        </div>

        {/* Actions */}
        <div className="flex items-center justify-end gap-2 px-4 py-3 bg-gray-50 dark:bg-[#1f1e1d] border-t border-gray-200 dark:border-[#3b3a39]">
          <button
            onClick={onClose}
            className="px-3 py-1.5 text-xs font-medium rounded hover:bg-gray-200 dark:hover:bg-[#323130] text-gray-600 dark:text-gray-300"
          >
            {t.cancel}
          </button>
          <button
            onClick={handleConfirm}
            className="px-4 py-1.5 text-xs font-medium rounded bg-[#185abd] hover:bg-[#104a9e] text-white"
          >
            {t.insert}
          </button>
        </div>
      </div>
    </div>
  );
};

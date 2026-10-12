import React, { useState } from 'react';
import { X, Calculator } from 'lucide-react';
import { Language } from '../../types/office';
import { getTranslation } from '../../i18n/translations';

interface FormulaHelperModalProps {
  isOpen: boolean;
  onClose: () => void;
  lang: Language;
  activeCellRef: string;
  onApplyFormula: (formula: string) => void;
}

interface FormulaDef {
  name: string;
  category: 'Math' | 'Statistical' | 'Logical' | 'Text';
  descEn: string;
  descAr: string;
  syntax: string;
  defaultTemplate: string;
}

const FORMULAS: FormulaDef[] = [
  {
    name: 'SUM',
    category: 'Math',
    descEn: 'Adds all numbers in a range of cells.',
    descAr: 'يجمع كافة الأرقام في نطاق محدد من الخلايا.',
    syntax: '=SUM(range)',
    defaultTemplate: '=SUM({range})',
  },
  {
    name: 'AVERAGE',
    category: 'Statistical',
    descEn: 'Calculates the arithmetic mean of a range of numbers.',
    descAr: 'يحسب المتوسط الحسابي لنطاق من الأرقام.',
    syntax: '=AVERAGE(range)',
    defaultTemplate: '=AVERAGE({range})',
  },
  {
    name: 'COUNT',
    category: 'Statistical',
    descEn: 'Counts how many cells contain numbers.',
    descAr: 'يحسب عدد الخلايا التي تحتوي على أرقام.',
    syntax: '=COUNT(range)',
    defaultTemplate: '=COUNT({range})',
  },
  {
    name: 'MIN',
    category: 'Statistical',
    descEn: 'Returns the minimum value in a set of numbers.',
    descAr: 'يستخرج أصغر قيمة رقمية في نطاق محدد.',
    syntax: '=MIN(range)',
    defaultTemplate: '=MIN({range})',
  },
  {
    name: 'MAX',
    category: 'Statistical',
    descEn: 'Returns the maximum value in a set of numbers.',
    descAr: 'يستخرج أكبر قيمة رقمية في نطاق محدد.',
    syntax: '=MAX(range)',
    defaultTemplate: '=MAX({range})',
  },
  {
    name: 'IF',
    category: 'Logical',
    descEn: 'Checks whether a condition is true, returning one value if true and another if false.',
    descAr: 'يتحقق من شرط منطقي، ويعيد قيمة إذا تحقق وقيمة أخرى إذا لم يتحقق.',
    syntax: '=IF(condition, value_if_true, value_if_false)',
    defaultTemplate: '=IF({range}>0, "Yes", "No")',
  },
  {
    name: 'ROUND',
    category: 'Math',
    descEn: 'Rounds a number to a specified number of digits.',
    descAr: 'يقرب الرقم إلى عدد محدد من المنازل العشرية.',
    syntax: '=ROUND(number, num_digits)',
    defaultTemplate: '=ROUND({range}, 2)',
  },
  {
    name: 'CONCAT',
    category: 'Text',
    descEn: 'Concatenates (joins) two or more text strings together.',
    descAr: 'يدمج سلسلتين نصيتين أو أكثر في نص واحد.',
    syntax: '=CONCAT(text1, text2, ...)',
    defaultTemplate: '=CONCAT({range}, " ")',
  },
];

export const FormulaHelperModal: React.FC<FormulaHelperModalProps> = ({
  isOpen,
  onClose,
  lang,
  activeCellRef,
  onApplyFormula,
}) => {
  const [selectedFunc, setSelectedFunc] = useState<FormulaDef>(FORMULAS[0]);
  const [customRange, setCustomRange] = useState('A1:A5');
  const t = getTranslation(lang);

  if (!isOpen) return null;

  const handleApply = () => {
    const formulaStr = selectedFunc.defaultTemplate.replace('{range}', customRange.trim().toUpperCase() || activeCellRef);
    onApplyFormula(formulaStr);
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/45 backdrop-blur-[1px] p-4">
      <div className="w-full max-w-lg rounded-lg bg-white dark:bg-[#252423] text-gray-800 dark:text-gray-100 shadow-2xl border border-gray-300 dark:border-[#3b3a39] overflow-hidden">
        {/* Header */}
        <div className="flex items-center justify-between px-4 py-3 bg-[#107c41] text-white">
          <div className="flex items-center gap-2">
            <Calculator className="w-4 h-4" />
            <h3 className="font-semibold text-sm">{t.formulaBuilderTitle}</h3>
          </div>
          <button onClick={onClose} className="p-1 rounded hover:bg-white/20">
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Content */}
        <div className="p-4 grid grid-cols-1 md:grid-cols-2 gap-4">
          {/* List of functions */}
          <div>
            <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1.5">
              {t.selectFunction}
            </label>
            <div className="border border-gray-200 dark:border-[#3b3a39] rounded h-56 overflow-y-auto divide-y divide-gray-100 dark:divide-[#323130] bg-gray-50 dark:bg-[#1b1a19]">
              {FORMULAS.map((fn) => (
                <button
                  key={fn.name}
                  onClick={() => setSelectedFunc(fn)}
                  className={`w-full text-left px-3 py-2 text-xs flex items-center justify-between transition-colors ${
                    selectedFunc.name === fn.name
                      ? 'bg-[#107c41] text-white font-medium'
                      : 'hover:bg-gray-200 dark:hover:bg-[#323130]'
                  }`}
                >
                  <span className="font-mono font-semibold">{fn.name}</span>
                  <span className={`text-[10px] px-1.5 py-0.5 rounded ${
                    selectedFunc.name === fn.name ? 'bg-white/20 text-white' : 'bg-gray-200 dark:bg-gray-700 text-gray-600 dark:text-gray-300'
                  }`}>
                    {fn.category}
                  </span>
                </button>
              ))}
            </div>
          </div>

          {/* Details & Target Range */}
          <div className="flex flex-col justify-between space-y-3">
            <div className="space-y-3">
              <div className="bg-gray-50 dark:bg-[#1b1a19] p-3 rounded border border-gray-200 dark:border-[#3b3a39]">
                <div className="text-xs font-semibold text-[#107c41] dark:text-green-400 font-mono mb-1">
                  {selectedFunc.syntax}
                </div>
                <div className="text-xs text-gray-600 dark:text-gray-300 leading-relaxed">
                  {lang === 'ar' ? selectedFunc.descAr : selectedFunc.descEn}
                </div>
              </div>

              <div>
                <label className="block text-xs font-medium text-gray-600 dark:text-gray-300 mb-1">
                  Target Range or Argument:
                </label>
                <input
                  type="text"
                  value={customRange}
                  onChange={(e) => setCustomRange(e.target.value)}
                  className="w-full px-3 py-1.5 text-xs font-mono rounded border border-gray-300 dark:border-[#484644] bg-white dark:bg-[#1b1a19] focus:outline-none focus:ring-2 focus:ring-[#107c41]"
                  placeholder="e.g. A1:A10"
                />
                <p className="text-[10px] text-gray-400 mt-1">
                  Applying to active cell: <strong className="text-[#107c41]">{activeCellRef}</strong>
                </p>
              </div>
            </div>

            <div className="p-2.5 bg-green-50 dark:bg-green-950/20 border border-green-200 dark:border-green-800 rounded text-xs font-mono text-green-800 dark:text-green-300">
              Preview: {selectedFunc.defaultTemplate.replace('{range}', customRange || activeCellRef)}
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="flex items-center justify-end gap-2 px-4 py-3 bg-gray-50 dark:bg-[#1f1e1d] border-t border-gray-200 dark:border-[#3b3a39]">
          <button
            onClick={onClose}
            className="px-3 py-1.5 text-xs font-medium rounded hover:bg-gray-200 dark:hover:bg-[#323130] text-gray-600 dark:text-gray-300"
          >
            {t.cancel}
          </button>
          <button
            onClick={handleApply}
            className="px-4 py-1.5 text-xs font-medium rounded bg-[#107c41] hover:bg-[#0b5c30] text-white"
          >
            {t.applyFormula}
          </button>
        </div>
      </div>
    </div>
  );
};

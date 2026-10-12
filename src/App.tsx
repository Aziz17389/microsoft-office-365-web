import React, { useState, useEffect, useRef, useCallback } from 'react';
import {
  AppMode,
  CellData,
  ExcelWorkbookState,
  FormattingState,
  GatedFeature,
  Language,
  PlanTier,
  RecentDocument,
  RibbonTab,
  Theme,
  WordEditorState,
  WordPageMargins,
  WordPageOrientation,
  WordPaperSize,
} from './types/office';
import { OfficeHeader } from './components/Header/OfficeHeader';
import { OfficeRibbon } from './components/Ribbon/OfficeRibbon';
import { WordHomeRibbon } from './components/Ribbon/WordHomeRibbon';
import { ExcelHomeRibbon } from './components/Ribbon/ExcelHomeRibbon';
import { InsertRibbon } from './components/Ribbon/InsertRibbon';
import { LayoutRibbon } from './components/Ribbon/LayoutRibbon';
import { ViewRibbon } from './components/Ribbon/ViewRibbon';
import { WordEditor } from './components/Word/WordEditor';
import { ExcelEditor } from './components/Excel/ExcelEditor';
import { OfficeLauncher } from './components/Launcher/OfficeLauncher';
import { FileMenuModal } from './components/Ribbon/FileMenuModal';
import { FindReplaceModal } from './components/Modals/FindReplaceModal';
import { InsertTableModal } from './components/Modals/InsertTableModal';
import { InsertLinkModal } from './components/Modals/InsertLinkModal';
import { InsertSymbolModal } from './components/Modals/InsertSymbolModal';
import { FormulaHelperModal } from './components/Modals/FormulaHelperModal';
import { PaywallModal } from './components/Modals/PaywallModal';
import {
  deleteRecentDocument,
  EXCEL_TEMPLATES,
  getRecentDocuments,
  loadExcelWorkbook,
  loadWordDocument,
  saveExcelWorkbook,
  saveWordDocument,
  WORD_TEMPLATES,
} from './services/storage';
import { indexToColName, parseCellRef } from './services/excelFormula';

export default function App() {
  // Global settings
  const [lang, setLang] = useState<Language>(() => {
    return (localStorage.getItem('office365_lang') as Language) || 'en';
  });
  const [theme, setTheme] = useState<Theme>(() => {
    return (localStorage.getItem('office365_theme') as Theme) || 'light';
  });

  // Subscription Plan & Paywall System
  const [plan, setPlan] = useState<PlanTier>(() => {
    return (localStorage.getItem('office365_plan') as PlanTier) || 'free';
  });
  const [secondsRemaining, setSecondsRemaining] = useState<number>(600); // 10 minutes = 600s
  const [isPaywallOpen, setIsPaywallOpen] = useState(false);
  const [targetedGatedFeature, setTargetedGatedFeature] = useState<GatedFeature | null>(null);

  const isTrialExpired = plan === 'free' && secondsRemaining <= 0;

  // Application routing mode
  const [appMode, setAppMode] = useState<AppMode>('launcher');
  const [activeRibbonTab, setActiveRibbonTab] = useState<RibbonTab>('home');
  const [recentDocs, setRecentDocs] = useState<RecentDocument[]>([]);
  const [isSaved, setIsSaved] = useState(true);

  // Modals state
  const [isFileMenuOpen, setIsFileMenuOpen] = useState(false);
  const [isFindReplaceOpen, setIsFindReplaceOpen] = useState(false);
  const [isTableModalOpen, setIsTableModalOpen] = useState(false);
  const [isLinkModalOpen, setIsLinkModalOpen] = useState(false);
  const [isSymbolModalOpen, setIsSymbolModalOpen] = useState(false);
  const [isFormulaHelperOpen, setIsFormulaHelperOpen] = useState(false);

  // Word Editor state
  const [wordState, setWordState] = useState<WordEditorState>({
    id: 'doc-' + Date.now(),
    title: 'Document 1',
    htmlContent: '<p>Start typing your document here...</p>',
    margins: 'normal',
    orientation: 'portrait',
    paperSize: 'A4',
    zoom: 100,
    showRuler: true,
    showGridlines: false,
    isZenMode: false,
  });

  const [wordFormatting, setWordFormatting] = useState<FormattingState>({
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    subscript: false,
    superscript: false,
    fontName: 'Segoe UI',
    fontSize: '11',
    foreColor: '#000000',
    hiliteColor: 'transparent',
    align: 'left',
  });

  const [isFormatPainterActive, setIsFormatPainterActive] = useState(false);
  const [savedFormatStyle, setSavedFormatStyle] = useState<FormattingState | null>(null);

  // Excel Workbook state
  const [excelState, setExcelState] = useState<ExcelWorkbookState>({
    id: 'book-' + Date.now(),
    title: 'Book 1',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: {},
        colWidths: {},
        rowHeights: {},
        rowCount: 45,
        colCount: 20,
      },
    ],
    activeSheetId: 'sheet-1',
    selectedCell: 'A1',
    selectionRange: null,
    zoom: 100,
    showGridlines: true,
    isZenMode: false,
  });

  const wordEditorRef = useRef<HTMLDivElement>(null);
  const autoSaveTimerRef = useRef<any>(null);

  // Initialize theme, language direction, and recent documents
  useEffect(() => {
    document.documentElement.setAttribute('dir', lang === 'ar' ? 'rtl' : 'ltr');
    document.documentElement.setAttribute('lang', lang);
    localStorage.setItem('office365_lang', lang);
  }, [lang]);

  useEffect(() => {
    if (theme === 'dark') {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
    localStorage.setItem('office365_theme', theme);
  }, [theme]);

  useEffect(() => {
    setRecentDocs(getRecentDocuments());
  }, [appMode]);

  // Handle escape to exit Zen mode
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        if (wordState.isZenMode) {
          setWordState((prev) => ({ ...prev, isZenMode: false }));
        }
        if (excelState.isZenMode) {
          setExcelState((prev) => ({ ...prev, isZenMode: false }));
        }
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [wordState.isZenMode, excelState.isZenMode]);

  // Theme & Language toggles
  const handleToggleTheme = () => {
    setTheme((prev) => (prev === 'light' ? 'dark' : 'light'));
  };

  const handleToggleLanguage = () => {
    setLang((prev) => (prev === 'en' ? 'ar' : 'en'));
  };

  // 10-Minute Countdown Timer for Free Plan
  useEffect(() => {
    if (plan === 'pro') return;
    if (secondsRemaining <= 0) {
      setIsPaywallOpen(true);
      return;
    }

    const timer = setInterval(() => {
      setSecondsRemaining((prev) => {
        if (prev <= 1) {
          clearInterval(timer);
          setIsPaywallOpen(true);
          return 0;
        }
        return prev - 1;
      });
    }, 1000);

    return () => clearInterval(timer);
  }, [plan, secondsRemaining]);

  // Subscription action handlers
  const handleTriggerGated = (feature: GatedFeature) => {
    setTargetedGatedFeature(feature);
    setIsPaywallOpen(true);
  };

  const handleUpgradeToPro = () => {
    setPlan('pro');
    localStorage.setItem('office365_plan', 'pro');
    setIsPaywallOpen(false);
  };

  const handleTogglePlan = (newPlan: PlanTier) => {
    setPlan(newPlan);
    localStorage.setItem('office365_plan', newPlan);
    if (newPlan === 'free') {
      setSecondsRemaining(600);
    }
  };

  const handleResetTimer = () => {
    setSecondsRemaining(600);
    if (plan === 'free') {
      setIsPaywallOpen(false);
    }
  };

  // Auto-save triggers (only for Pro plan; Free plan has saving restricted)
  const triggerAutoSave = useCallback(() => {
    if (plan !== 'pro') {
      setIsSaved(false);
      return;
    }
    setIsSaved(false);
    if (autoSaveTimerRef.current) {
      clearTimeout(autoSaveTimerRef.current);
    }
    autoSaveTimerRef.current = setTimeout(() => {
      if (appMode === 'word') {
        saveWordDocument(wordState);
      } else if (appMode === 'excel') {
        saveExcelWorkbook(excelState);
      }
      setIsSaved(true);
      setRecentDocs(getRecentDocuments());
    }, 600);
  }, [appMode, wordState, excelState, plan]);

  // Execute manual save (blocked in Free plan)
  const handleManualSave = () => {
    if (plan !== 'pro') {
      handleTriggerGated('save_export');
      return;
    }
    if (appMode === 'word') {
      saveWordDocument(wordState);
    } else if (appMode === 'excel') {
      saveExcelWorkbook(excelState);
    }
    setIsSaved(true);
    setRecentDocs(getRecentDocuments());
  };

  // Open Word template
  const handleOpenWordTemplate = (templateKey: string) => {
    const tpl = WORD_TEMPLATES[templateKey] || WORD_TEMPLATES.blank;
    const newDoc: WordEditorState = {
      id: 'doc-' + Date.now(),
      title: tpl.title,
      htmlContent: tpl.html,
      margins: 'normal',
      orientation: 'portrait',
      paperSize: 'A4',
      zoom: 100,
      showRuler: true,
      showGridlines: false,
      isZenMode: false,
    };
    setWordState(newDoc);
    saveWordDocument(newDoc);
    setAppMode('word');
    setActiveRibbonTab('home');
  };

  // Open Excel template
  const handleOpenExcelTemplate = (templateKey: string) => {
    const tpl = EXCEL_TEMPLATES[templateKey] || EXCEL_TEMPLATES.blank;
    const newBook: ExcelWorkbookState = {
      id: 'book-' + Date.now(),
      title: tpl.title,
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet1',
          cells: tpl.cells,
          colWidths: { A: 160, B: 110, C: 110, D: 110 },
          rowHeights: {},
          rowCount: 45,
          colCount: 20,
        },
      ],
      activeSheetId: 'sheet-1',
      selectedCell: 'A1',
      selectionRange: null,
      zoom: 100,
      showGridlines: true,
      isZenMode: false,
    };
    setExcelState(newBook);
    saveExcelWorkbook(newBook);
    setAppMode('excel');
    setActiveRibbonTab('home');
  };

  // Open from recent documents
  const handleOpenRecentDoc = (doc: RecentDocument) => {
    if (doc.app === 'word') {
      const loaded = loadWordDocument(doc.id);
      if (loaded) {
        setWordState(loaded);
        setAppMode('word');
        setActiveRibbonTab('home');
      }
    } else if (doc.app === 'excel') {
      const loaded = loadExcelWorkbook(doc.id);
      if (loaded) {
        setExcelState(loaded);
        setAppMode('excel');
        setActiveRibbonTab('home');
      }
    }
  };

  // Delete recent document
  const handleDeleteRecentDoc = (docId: string) => {
    deleteRecentDocument(docId);
    setRecentDocs(getRecentDocuments());
  };

  // Import file (.docx, .xlsx, .csv, .txt, .html)
  const handleImportFile = async (file: File) => {
    const fileName = file.name.replace(/\.[^/.]+$/, '');
    const extension = file.name.split('.').pop()?.toLowerCase();

    if (extension === 'csv') {
      const text = await file.text();
      const lines = text.split(/\r?\n/);
      const cells: Record<string, CellData> = {};

      lines.forEach((line, rIdx) => {
        const cols = line.split(',');
        cols.forEach((val, cIdx) => {
          const colLetter = indexToColName(cIdx);
          const cellKey = `${colLetter}${rIdx + 1}`;
          cells[cellKey] = { raw: val.trim() };
        });
      });

      const newBook: ExcelWorkbookState = {
        id: 'book-' + Date.now(),
        title: fileName,
        sheets: [
          {
            id: 'sheet-1',
            name: 'Sheet1',
            cells,
            colWidths: {},
            rowHeights: {},
            rowCount: Math.max(lines.length + 10, 40),
            colCount: 25,
          },
        ],
        activeSheetId: 'sheet-1',
        selectedCell: 'A1',
        selectionRange: null,
        zoom: 100,
        showGridlines: true,
        isZenMode: false,
      };
      setExcelState(newBook);
      saveExcelWorkbook(newBook);
      setAppMode('excel');
    } else if (extension === 'json') {
      try {
        const text = await file.text();
        const parsed = JSON.parse(text);
        if (parsed.sheets) {
          setExcelState(parsed);
          saveExcelWorkbook(parsed);
          setAppMode('excel');
        }
      } catch (err) {
        console.error('Invalid JSON', err);
      }
    } else {
      // Word text or html
      const text = await file.text();
      const htmlContent = extension === 'html' ? text : `<pre style="white-space: pre-wrap;">${text}</pre>`;
      const newDoc: WordEditorState = {
        id: 'doc-' + Date.now(),
        title: fileName,
        htmlContent,
        margins: 'normal',
        orientation: 'portrait',
        paperSize: 'A4',
        zoom: 100,
        showRuler: true,
        showGridlines: false,
        isZenMode: false,
      };
      setWordState(newDoc);
      saveWordDocument(newDoc);
      setAppMode('word');
    }
  };

  // Word content updater
  const handleWordContentChange = (html: string) => {
    setWordState((prev) => {
      const updated = { ...prev, htmlContent: html };
      return updated;
    });
    triggerAutoSave();
  };

  // Word Ribbon Commands
  const handleWordCommand = (command: string, value: any = null) => {
    if (wordEditorRef.current) {
      wordEditorRef.current.focus();
    }
    if (command === 'fontSize') {
      // Map pt to execCommand font size (1 to 7)
      const pt = parseInt(value, 10);
      let sizeIdx = '3';
      if (pt <= 9) sizeIdx = '1';
      else if (pt <= 10) sizeIdx = '2';
      else if (pt <= 12) sizeIdx = '3';
      else if (pt <= 14) sizeIdx = '4';
      else if (pt <= 18) sizeIdx = '5';
      else if (pt <= 24) sizeIdx = '6';
      else sizeIdx = '7';
      document.execCommand('fontSize', false, sizeIdx);
    } else if (command === 'lineSpacing') {
      if (wordEditorRef.current) {
        const sel = window.getSelection();
        if (sel && sel.rangeCount > 0) {
          const parent = sel.anchorNode?.parentElement;
          if (parent) {
            parent.style.lineHeight = value;
          }
        }
      }
    } else {
      document.execCommand(command, false, value);
    }

    if (wordEditorRef.current) {
      handleWordContentChange(wordEditorRef.current.innerHTML);
    }
  };

  // Format painter toggle
  const handleToggleFormatPainter = () => {
    if (!isFormatPainterActive) {
      setSavedFormatStyle({ ...wordFormatting });
      setIsFormatPainterActive(true);
    } else {
      setIsFormatPainterActive(false);
      setSavedFormatStyle(null);
    }
  };

  // Insert image handler
  const handleInsertImage = (dataUrl: string) => {
    if (appMode === 'word') {
      const imgHtml = `<img src="${dataUrl}" style="max-width: 100%; height: auto; border-radius: 4px; margin: 12px 0;" alt="Inserted image" />`;
      document.execCommand('insertHTML', false, imgHtml);
      if (wordEditorRef.current) {
        handleWordContentChange(wordEditorRef.current.innerHTML);
      }
    }
  };

  // Insert Table handler
  const handleInsertTable = (rows: number, cols: number) => {
    if (appMode === 'word') {
      let tableHtml = `<table style="width: 100%; border-collapse: collapse; margin: 16px 0; border: 1px solid #d1d5db;"><tbody>`;
      for (let r = 0; r < rows; r++) {
        tableHtml += '<tr>';
        for (let c = 0; c < cols; c++) {
          const isHeader = r === 0;
          const bg = isHeader ? 'background-color: #f3f2f1; font-weight: bold;' : '';
          tableHtml += `<td style="border: 1px solid #d1d5db; padding: 8px; ${bg}"><p><br></p></td>`;
        }
        tableHtml += '</tr>';
      }
      tableHtml += '</tbody></table><p><br></p>';
      document.execCommand('insertHTML', false, tableHtml);
      if (wordEditorRef.current) {
        handleWordContentChange(wordEditorRef.current.innerHTML);
      }
    } else if (appMode === 'excel') {
      // In Excel, populate active cell range with sample table headers
      const parsed = parseCellRef(excelState.selectedCell) || { colIdx: 0, row: 1 };
      const activeSheet = excelState.sheets.find((s) => s.id === excelState.activeSheetId);
      if (!activeSheet) return;

      const newCells = { ...activeSheet.cells };
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const cName = indexToColName(parsed.colIdx + c);
          const rNum = parsed.row + r;
          const ref = `${cName}${rNum}`;
          if (r === 0) {
            newCells[ref] = { raw: `Header ${c + 1}`, bold: true, bgColor: '#dff6dd' };
          } else {
            newCells[ref] = { raw: `Data ${r},${c + 1}` };
          }
        }
      }
      setExcelState((prev) => ({
        ...prev,
        sheets: prev.sheets.map((s) => (s.id === prev.activeSheetId ? { ...s, cells: newCells } : s)),
      }));
      triggerAutoSave();
    }
  };

  // Insert Link handler
  const handleInsertLink = (text: string, url: string) => {
    if (appMode === 'word') {
      const linkHtml = `<a href="${url}" target="_blank" rel="noopener noreferrer" style="color: #185abd; text-decoration: underline;">${text}</a>`;
      document.execCommand('insertHTML', false, linkHtml);
      if (wordEditorRef.current) {
        handleWordContentChange(wordEditorRef.current.innerHTML);
      }
    }
  };

  // Insert Symbol handler
  const handleInsertSymbol = (symbol: string) => {
    if (appMode === 'word') {
      document.execCommand('insertText', false, symbol);
      if (wordEditorRef.current) {
        handleWordContentChange(wordEditorRef.current.innerHTML);
      }
    } else if (appMode === 'excel') {
      const activeSheet = excelState.sheets.find((s) => s.id === excelState.activeSheetId);
      if (activeSheet) {
        const cur = activeSheet.cells[excelState.selectedCell]?.raw || '';
        handleUpdateExcelCell(activeSheet.id, excelState.selectedCell, { raw: cur + symbol });
      }
    }
  };

  // Excel Cell update
  const handleUpdateExcelCell = (sheetId: string, cellRef: string, data: Partial<CellData>) => {
    setExcelState((prev) => {
      const updatedSheets = prev.sheets.map((s) => {
        if (s.id !== sheetId) return s;
        const currentCell = s.cells[cellRef] || { raw: '' };
        return {
          ...s,
          cells: {
            ...s.cells,
            [cellRef]: { ...currentCell, ...data },
          },
        };
      });
      return { ...prev, sheets: updatedSheets };
    });
    triggerAutoSave();
  };

  // Excel Cell selection
  const handleExcelSelectCell = (cellRef: string, range?: { start: string; end: string } | null) => {
    setExcelState((prev) => ({
      ...prev,
      selectedCell: cellRef,
      selectionRange: range !== undefined ? range : null,
    }));
  };

  // Excel Sheet actions
  const handleExcelAddSheet = () => {
    setExcelState((prev) => {
      const newSheetNumber = prev.sheets.length + 1;
      const newSheetId = 'sheet-' + Date.now();
      const newSheet = {
        id: newSheetId,
        name: `Sheet${newSheetNumber}`,
        cells: {},
        colWidths: {},
        rowHeights: {},
        rowCount: 45,
        colCount: 20,
      };
      return {
        ...prev,
        sheets: [...prev.sheets, newSheet],
        activeSheetId: newSheetId,
      };
    });
    triggerAutoSave();
  };

  const handleExcelRenameSheet = (sheetId: string, newName: string) => {
    setExcelState((prev) => ({
      ...prev,
      sheets: prev.sheets.map((s) => (s.id === sheetId ? { ...s, name: newName } : s)),
    }));
    triggerAutoSave();
  };

  const handleExcelDeleteSheet = (sheetId: string) => {
    setExcelState((prev) => {
      if (prev.sheets.length <= 1) return prev;
      const filtered = prev.sheets.filter((s) => s.id !== sheetId);
      return {
        ...prev,
        sheets: filtered,
        activeSheetId: prev.activeSheetId === sheetId ? filtered[0].id : prev.activeSheetId,
      };
    });
    triggerAutoSave();
  };

  // Excel AutoSum (∑) action
  const handleExcelAutoSum = () => {
    const activeSheet = excelState.sheets.find((s) => s.id === excelState.activeSheetId);
    if (!activeSheet) return;

    const parsed = parseCellRef(excelState.selectedCell);
    if (!parsed) return;

    // Check cells directly above
    const aboveRow = parsed.row - 1;
    if (aboveRow >= 1) {
      const sumFormula = `=SUM(${parsed.col}1:${parsed.col}${aboveRow})`;
      handleUpdateExcelCell(activeSheet.id, excelState.selectedCell, {
        raw: sumFormula,
        bold: true,
      });
    }
  };

  // Excel Sort action
  const handleExcelSort = (ascending: boolean) => {
    const activeSheet = excelState.sheets.find((s) => s.id === excelState.activeSheetId);
    if (!activeSheet) return;

    const parsed = parseCellRef(excelState.selectedCell);
    if (!parsed) return;

    const col = parsed.col;
    // Extract row values for this column
    const rowValues: { row: number; val: string | number }[] = [];
    for (let r = 1; r <= activeSheet.rowCount; r++) {
      const cell = activeSheet.cells[`${col}${r}`];
      if (cell && cell.raw) {
        rowValues.push({ row: r, val: cell.raw });
      }
    }

    rowValues.sort((a, b) => {
      const numA = Number(a.val);
      const numB = Number(b.val);
      if (!isNaN(numA) && !isNaN(numB)) {
        return ascending ? numA - numB : numB - numA;
      }
      return ascending
        ? String(a.val).localeCompare(String(b.val))
        : String(b.val).localeCompare(String(a.val));
    });

    const newCells = { ...activeSheet.cells };
    rowValues.forEach((item, idx) => {
      const targetRow = idx + 1;
      newCells[`${col}${targetRow}`] = {
        ...newCells[`${col}${targetRow}`],
        raw: String(item.val),
      };
    });

    setExcelState((prev) => ({
      ...prev,
      sheets: prev.sheets.map((s) => (s.id === prev.activeSheetId ? { ...s, cells: newCells } : s)),
    }));
    triggerAutoSave();
  };

  // Excel Clear action
  const handleExcelClear = () => {
    const activeSheet = excelState.sheets.find((s) => s.id === excelState.activeSheetId);
    if (!activeSheet) return;

    handleUpdateExcelCell(activeSheet.id, excelState.selectedCell, { raw: '' });
  };

  // Find & Replace actions
  const handleFind = (search: string, matchCase: boolean) => {
    if (appMode === 'word') {
      if (typeof (window as any).find === 'function') {
        (window as any).find(search, matchCase, false, true);
      }
    }
  };

  const handleReplace = (search: string, replaceWith: string, matchCase: boolean) => {
    if (appMode === 'word' && wordEditorRef.current) {
      const content = wordEditorRef.current.innerHTML;
      const flags = matchCase ? 'g' : 'gi';
      const regex = new RegExp(search.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags);
      const updated = content.replace(regex, replaceWith);
      wordEditorRef.current.innerHTML = updated;
      handleWordContentChange(updated);
    } else if (appMode === 'excel') {
      const activeSheet = excelState.sheets.find((s) => s.id === excelState.activeSheetId);
      if (!activeSheet) return;
      const cellData = activeSheet.cells[excelState.selectedCell];
      if (cellData && cellData.raw) {
        const flags = matchCase ? '' : 'i';
        const regex = new RegExp(search.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags);
        const newRaw = cellData.raw.replace(regex, replaceWith);
        handleUpdateExcelCell(activeSheet.id, excelState.selectedCell, { raw: newRaw });
      }
    }
  };

  const handleReplaceAll = (search: string, replaceWith: string, matchCase: boolean) => {
    if (appMode === 'word' && wordEditorRef.current) {
      const content = wordEditorRef.current.innerHTML;
      const flags = matchCase ? 'g' : 'gi';
      const regex = new RegExp(search.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags);
      const updated = content.replace(regex, replaceWith);
      wordEditorRef.current.innerHTML = updated;
      handleWordContentChange(updated);
    } else if (appMode === 'excel') {
      const activeSheet = excelState.sheets.find((s) => s.id === excelState.activeSheetId);
      if (!activeSheet) return;
      const flags = matchCase ? 'g' : 'gi';
      const regex = new RegExp(search.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags);
      const newCells = { ...activeSheet.cells };

      Object.keys(newCells).forEach((key) => {
        if (newCells[key]?.raw) {
          newCells[key] = {
            ...newCells[key],
            raw: newCells[key].raw.replace(regex, replaceWith),
          };
        }
      });

      setExcelState((prev) => ({
        ...prev,
        sheets: prev.sheets.map((s) => (s.id === prev.activeSheetId ? { ...s, cells: newCells } : s)),
      }));
      triggerAutoSave();
    }
  };

  // Export handlers
  const handleExport = (format: string) => {
    if (appMode === 'word') {
      const title = wordState.title || 'Document';
      if (format === 'html' || format === 'doc') {
        const fullHtml = `
          <!DOCTYPE html>
          <html lang="${lang}" dir="${lang === 'ar' ? 'rtl' : 'ltr'}">
          <head>
            <meta charset="utf-8">
            <title>${title}</title>
            <style>
              body { font-family: Segoe UI, Arial, sans-serif; padding: 40px; max-width: 800px; margin: auto; }
              table { border-collapse: collapse; width: 100%; }
              td, th { border: 1px solid #ccc; padding: 8px; }
            </style>
          </head>
          <body>
            ${wordState.htmlContent}
          </body>
          </html>
        `;
        const blob = new Blob([fullHtml], { type: format === 'doc' ? 'application/msword' : 'text/html' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${title}.${format === 'doc' ? 'doc' : 'html'}`;
        a.click();
        URL.revokeObjectURL(url);
      } else if (format === 'txt') {
        const temp = document.createElement('div');
        temp.innerHTML = wordState.htmlContent;
        const text = temp.innerText || temp.textContent || '';
        const blob = new Blob([text], { type: 'text/plain' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${title}.txt`;
        a.click();
        URL.revokeObjectURL(url);
      }
    } else if (appMode === 'excel') {
      const activeSheet = excelState.sheets.find((s) => s.id === excelState.activeSheetId) || excelState.sheets[0];
      const title = excelState.title || 'Workbook';

      if (format === 'csv') {
        let csvContent = '';
        for (let r = 1; r <= activeSheet.rowCount; r++) {
          const rowVals: string[] = [];
          for (let c = 0; c < activeSheet.colCount; c++) {
            const col = indexToColName(c);
            const val = activeSheet.cells[`${col}${r}`]?.raw || '';
            rowVals.push(`"${val.replace(/"/g, '""')}"`);
          }
          csvContent += rowVals.join(',') + '\r\n';
        }
        const blob = new Blob([csvContent], { type: 'text/csv' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${title}.csv`;
        a.click();
        URL.revokeObjectURL(url);
      } else if (format === 'json') {
        const jsonStr = JSON.stringify(excelState, null, 2);
        const blob = new Blob([jsonStr], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${title}.json`;
        a.click();
        URL.revokeObjectURL(url);
      } else if (format === 'html') {
        let tableHtml = '<table border="1" cellpadding="6" style="border-collapse: collapse; font-family: Segoe UI, sans-serif;">';
        for (let r = 1; r <= activeSheet.rowCount; r++) {
          tableHtml += '<tr>';
          for (let c = 0; c < activeSheet.colCount; c++) {
            const col = indexToColName(c);
            const val = activeSheet.cells[`${col}${r}`]?.raw || '';
            tableHtml += `<td>${val}</td>`;
          }
          tableHtml += '</tr>';
        }
        tableHtml += '</table>';
        const blob = new Blob([tableHtml], { type: 'text/html' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${title}.html`;
        a.click();
        URL.revokeObjectURL(url);
      }
    }
  };

  const handlePrint = () => {
    window.print();
  };

  // Quick Action execution from search bar
  const handleExecuteCommand = (cmd: string) => {
    switch (cmd) {
      case 'bold':
        if (appMode === 'word') handleWordCommand('bold');
        break;
      case 'table':
        setIsTableModalOpen(true);
        break;
      case 'find':
        setIsFindReplaceOpen(true);
        break;
      case 'print':
        handlePrint();
        break;
      case 'autosum':
        if (appMode === 'excel') handleExcelAutoSum();
        break;
      case 'currency':
        if (appMode === 'excel') {
          const activeSheet = excelState.sheets.find((s) => s.id === excelState.activeSheetId);
          if (activeSheet) {
            handleUpdateExcelCell(activeSheet.id, excelState.selectedCell, {
              format: 'currency',
              align: 'right',
            });
          }
        }
        break;
    }
  };

  const activeDocTitle = appMode === 'word' ? wordState.title : excelState.title;
  const isZenMode = appMode === 'word' ? wordState.isZenMode : excelState.isZenMode;

  return (
    <div className={`min-h-screen h-screen flex flex-col overflow-hidden bg-[#f3f2f1] dark:bg-[#11100f] text-gray-800 dark:text-gray-100 font-sans ${theme === 'dark' ? 'dark' : ''}`}>
      {/* 1. START SCREEN / LAUNCHER */}
      {appMode === 'launcher' ? (
        <OfficeLauncher
          lang={lang}
          theme={theme}
          plan={plan}
          secondsRemaining={secondsRemaining}
          onOpenUpgradeModal={() => handleTriggerGated('unlimited_time')}
          onTogglePlan={handleTogglePlan}
          recentDocs={recentDocs}
          onOpenWordTemplate={handleOpenWordTemplate}
          onOpenExcelTemplate={handleOpenExcelTemplate}
          onOpenRecentDoc={handleOpenRecentDoc}
          onDeleteRecentDoc={handleDeleteRecentDoc}
          onImportFile={handleImportFile}
          onToggleTheme={handleToggleTheme}
          onToggleLanguage={handleToggleLanguage}
        />
      ) : (
        /* 2. ACTIVE APPLICATION (WORD OR EXCEL) */
        <div className="flex-1 flex flex-col h-full overflow-hidden">
          {/* Header Bar */}
          {!isZenMode && (
            <OfficeHeader
              appMode={appMode}
              lang={lang}
              theme={theme}
              docTitle={activeDocTitle}
              isSaved={isSaved}
              plan={plan}
              secondsRemaining={secondsRemaining}
              onOpenUpgradeModal={() => handleTriggerGated('unlimited_time')}
              onTogglePlan={handleTogglePlan}
              onTitleChange={(newTitle) => {
                if (appMode === 'word') {
                  setWordState((p) => ({ ...p, title: newTitle }));
                } else {
                  setExcelState((p) => ({ ...p, title: newTitle }));
                }
                triggerAutoSave();
              }}
              onSave={handleManualSave}
              onUndo={() => {
                if (appMode === 'word') handleWordCommand('undo');
              }}
              onRedo={() => {
                if (appMode === 'word') handleWordCommand('redo');
              }}
              onSwitchApp={(newMode) => {
                setAppMode(newMode);
                if (newMode !== 'launcher') setActiveRibbonTab('home');
              }}
              onToggleTheme={handleToggleTheme}
              onToggleLanguage={handleToggleLanguage}
              onToggleZenMode={() => {
                if (appMode === 'word') {
                  setWordState((p) => ({ ...p, isZenMode: !p.isZenMode }));
                } else {
                  setExcelState((p) => ({ ...p, isZenMode: !p.isZenMode }));
                }
              }}
              onExecuteCommand={handleExecuteCommand}
            />
          )}

          {/* Ribbon Toolbar */}
          {!isZenMode && (
            <OfficeRibbon
              appMode={appMode}
              lang={lang}
              activeTab={activeRibbonTab}
              onSelectTab={setActiveRibbonTab}
              onOpenFileMenu={() => setIsFileMenuOpen(true)}
            >
              {/* Home Tab */}
              {activeRibbonTab === 'home' && (
                appMode === 'word' ? (
                  <WordHomeRibbon
                    lang={lang}
                    formatting={wordFormatting}
                    plan={plan}
                    onTriggerGated={handleTriggerGated}
                    onCommand={handleWordCommand}
                    onOpenFindReplace={() => setIsFindReplaceOpen(true)}
                    isFormatPainterActive={isFormatPainterActive}
                    onToggleFormatPainter={handleToggleFormatPainter}
                  />
                ) : (
                  <ExcelHomeRibbon
                    lang={lang}
                    plan={plan}
                    onTriggerGated={handleTriggerGated}
                    activeCellData={
                      excelState.sheets.find((s) => s.id === excelState.activeSheetId)?.cells[
                        excelState.selectedCell
                      ]
                    }
                    onUpdateFormat={(updates) => {
                      const activeSheet = excelState.sheets.find(
                        (s) => s.id === excelState.activeSheetId
                      );
                      if (activeSheet) {
                        handleUpdateExcelCell(activeSheet.id, excelState.selectedCell, updates);
                      }
                    }}
                    onAutoSum={handleExcelAutoSum}
                    onSort={handleExcelSort}
                    onClear={handleExcelClear}
                    onOpenFindReplace={() => setIsFindReplaceOpen(true)}
                    onOpenFormulaHelper={() => setIsFormulaHelperOpen(true)}
                  />
                )
              )}

              {/* Insert Tab */}
              {activeRibbonTab === 'insert' && (
                <InsertRibbon
                  appMode={appMode}
                  lang={lang}
                  plan={plan}
                  onTriggerGated={handleTriggerGated}
                  onOpenTableModal={() => setIsTableModalOpen(true)}
                  onOpenLinkModal={() => setIsLinkModalOpen(true)}
                  onOpenSymbolModal={() => setIsSymbolModalOpen(true)}
                  onInsertImage={handleInsertImage}
                  onInsertHorizontalRule={() => handleWordCommand('insertHorizontalRule')}
                  onInsertPageBreak={() => {
                    handleWordCommand('insertHTML', '<div style="page-break-before: always; border-top: 1px dashed #ccc; margin: 20px 0;"></div><p><br></p>');
                  }}
                />
              )}

              {/* Page Layout Tab */}
              {activeRibbonTab === 'layout' && (
                <LayoutRibbon
                  lang={lang}
                  plan={plan}
                  onTriggerGated={handleTriggerGated}
                  margins={wordState.margins}
                  orientation={wordState.orientation}
                  paperSize={wordState.paperSize}
                  onChangeMargins={(m: WordPageMargins) => {
                    setWordState((p) => ({ ...p, margins: m }));
                    triggerAutoSave();
                  }}
                  onChangeOrientation={(o: WordPageOrientation) => {
                    setWordState((p) => ({ ...p, orientation: o }));
                    triggerAutoSave();
                  }}
                  onChangePaperSize={(s: WordPaperSize) => {
                    setWordState((p) => ({ ...p, paperSize: s }));
                    triggerAutoSave();
                  }}
                />
              )}

              {/* View Tab */}
              {activeRibbonTab === 'view' && (
                <ViewRibbon
                  lang={lang}
                  showRuler={appMode === 'word' ? wordState.showRuler : false}
                  showGridlines={appMode === 'word' ? wordState.showGridlines : excelState.showGridlines}
                  isZenMode={isZenMode}
                  zoom={appMode === 'word' ? wordState.zoom : excelState.zoom}
                  onToggleRuler={() => {
                    setWordState((p) => ({ ...p, showRuler: !p.showRuler }));
                    triggerAutoSave();
                  }}
                  onToggleGridlines={() => {
                    if (appMode === 'word') {
                      setWordState((p) => ({ ...p, showGridlines: !p.showGridlines }));
                    } else {
                      setExcelState((p) => ({ ...p, showGridlines: !p.showGridlines }));
                    }
                    triggerAutoSave();
                  }}
                  onToggleZenMode={() => {
                    if (appMode === 'word') {
                      setWordState((p) => ({ ...p, isZenMode: !p.isZenMode }));
                    } else {
                      setExcelState((p) => ({ ...p, isZenMode: !p.isZenMode }));
                    }
                  }}
                  onSetZoom={(z) => {
                    if (appMode === 'word') {
                      setWordState((p) => ({ ...p, zoom: z }));
                    } else {
                      setExcelState((p) => ({ ...p, zoom: z }));
                    }
                  }}
                />
              )}
            </OfficeRibbon>
          )}

          {/* Zen Mode Exit button indicator */}
          {isZenMode && (
            <div className="fixed top-3 right-4 z-50">
              <button
                onClick={() => {
                  if (appMode === 'word') setWordState((p) => ({ ...p, isZenMode: false }));
                  else setExcelState((p) => ({ ...p, isZenMode: false }));
                }}
                className="bg-black/80 hover:bg-black text-white text-xs px-3 py-1.5 rounded-full shadow-lg backdrop-blur-sm transition-all"
              >
                Exit Fullscreen (Esc)
              </button>
            </div>
          )}

          {/* Core Document Engine (Word or Excel) */}
          <div className="flex-1 flex flex-col h-full overflow-hidden">
            {appMode === 'word' ? (
              <WordEditor
                lang={lang}
                state={wordState}
                editorRefProp={wordEditorRef}
                onChangeContent={handleWordContentChange}
                onUpdateFormatting={setWordFormatting}
                isFormatPainterActive={isFormatPainterActive}
                onDeactivateFormatPainter={() => {
                  setIsFormatPainterActive(false);
                  setSavedFormatStyle(null);
                }}
                savedFormatPainterStyle={savedFormatStyle}
                onSetZoom={(z) => setWordState((p) => ({ ...p, zoom: z }))}
              />
            ) : (
              <ExcelEditor
                lang={lang}
                workbook={excelState}
                onUpdateCell={handleUpdateExcelCell}
                onSelectCell={handleExcelSelectCell}
                onAddSheet={handleExcelAddSheet}
                onRenameSheet={handleExcelRenameSheet}
                onDeleteSheet={handleExcelDeleteSheet}
                onSetActiveSheet={(sheetId) => setExcelState((p) => ({ ...p, activeSheetId: sheetId }))}
                onSetZoom={(z) => setExcelState((p) => ({ ...p, zoom: z }))}
                onOpenFormulaHelper={() => setIsFormulaHelperOpen(true)}
              />
            )}
          </div>
        </div>
      )}

      {/* MODALS */}
      {/* 1. Backstage File Menu */}
      <FileMenuModal
        isOpen={isFileMenuOpen}
        onClose={() => setIsFileMenuOpen(false)}
        appMode={appMode}
        lang={lang}
        plan={plan}
        onTriggerGated={handleTriggerGated}
        docTitle={activeDocTitle}
        onNew={() => {
          if (appMode === 'word') handleOpenWordTemplate('blank');
          else handleOpenExcelTemplate('blank');
        }}
        onSave={handleManualSave}
        onSaveAs={(newTitle) => {
          if (appMode === 'word') {
            const copy = { ...wordState, id: 'doc-' + Date.now(), title: newTitle };
            setWordState(copy);
            saveWordDocument(copy);
          } else {
            const copy = { ...excelState, id: 'book-' + Date.now(), title: newTitle };
            setExcelState(copy);
            saveExcelWorkbook(copy);
          }
          setRecentDocs(getRecentDocuments());
        }}
        onExport={handleExport}
        onPrint={handlePrint}
        onCloseToHome={() => {
          setIsFileMenuOpen(false);
          setAppMode('launcher');
        }}
        wordCount={wordState.htmlContent.replace(/<[^>]*>/g, ' ').trim().split(/\s+/).length}
        cellCount={
          Object.keys(
            excelState.sheets.find((s) => s.id === excelState.activeSheetId)?.cells || {}
          ).length
        }
      />

      {/* 2. Find & Replace Modal */}
      <FindReplaceModal
        isOpen={isFindReplaceOpen}
        onClose={() => setIsFindReplaceOpen(false)}
        lang={lang}
        onFind={handleFind}
        onReplace={handleReplace}
        onReplaceAll={handleReplaceAll}
      />

      {/* 3. Insert Table Modal */}
      <InsertTableModal
        isOpen={isTableModalOpen}
        onClose={() => setIsTableModalOpen(false)}
        lang={lang}
        onInsertTable={handleInsertTable}
      />

      {/* 4. Insert Link Modal */}
      <InsertLinkModal
        isOpen={isLinkModalOpen}
        onClose={() => setIsLinkModalOpen(false)}
        lang={lang}
        onInsertLink={handleInsertLink}
      />

      {/* 5. Insert Symbol & Emoji Modal */}
      <InsertSymbolModal
        isOpen={isSymbolModalOpen}
        onClose={() => setIsSymbolModalOpen(false)}
        lang={lang}
        onInsertSymbol={handleInsertSymbol}
      />

      {/* 6. Excel Formula Builder Modal */}
      <FormulaHelperModal
        isOpen={isFormulaHelperOpen}
        onClose={() => setIsFormulaHelperOpen(false)}
        lang={lang}
        activeCellRef={excelState.selectedCell}
        onApplyFormula={(formula) => {
          const activeSheet = excelState.sheets.find((s) => s.id === excelState.activeSheetId);
          if (activeSheet) {
            handleUpdateExcelCell(activeSheet.id, excelState.selectedCell, { raw: formula });
          }
        }}
      />

      {/* 7. Office 365+ Subscription & Paywall Modal */}
      <PaywallModal
        isOpen={isPaywallOpen || isTrialExpired}
        onClose={() => {
          if (!isTrialExpired) {
            setIsPaywallOpen(false);
            setTargetedGatedFeature(null);
          }
        }}
        isTrialExpired={isTrialExpired}
        lang={lang}
        gatedFeature={targetedGatedFeature}
        onUpgrade={handleUpgradeToPro}
        currentPlan={plan}
        onTogglePlan={handleTogglePlan}
        onResetTimer={handleResetTimer}
      />
    </div>
  );
}

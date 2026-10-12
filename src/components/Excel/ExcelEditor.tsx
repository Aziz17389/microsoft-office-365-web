import React, { useState, useRef, useEffect, useCallback, useMemo } from 'react';
import {
  ExcelSheet,
  ExcelWorkbookState,
  Language,
  CellData,
} from '../../types/office';
import { getTranslation } from '../../i18n/translations';
import {
  calculateRangeStats,
  colNameToIndex,
  evaluateCell,
  expandCellRange,
  formatCellValue,
  indexToColName,
  parseCellRef,
} from '../../services/excelFormula';
import { Plus, X, ZoomIn, ZoomOut, Check, ChevronRight } from 'lucide-react';

interface ExcelEditorProps {
  lang: Language;
  workbook: ExcelWorkbookState;
  onUpdateCell: (sheetId: string, cellRef: string, data: Partial<CellData>) => void;
  onSelectCell: (cellRef: string, range?: { start: string; end: string } | null) => void;
  onAddSheet: () => void;
  onRenameSheet: (sheetId: string, newName: string) => void;
  onDeleteSheet: (sheetId: string) => void;
  onSetActiveSheet: (sheetId: string) => void;
  onSetZoom: (zoom: number) => void;
  onOpenFormulaHelper: () => void;
}

export const ExcelEditor: React.FC<ExcelEditorProps> = ({
  lang,
  workbook,
  onUpdateCell,
  onSelectCell,
  onAddSheet,
  onRenameSheet,
  onDeleteSheet,
  onSetActiveSheet,
  onSetZoom,
  onOpenFormulaHelper,
}) => {
  const activeSheet = useMemo(() => {
    return workbook.sheets.find((s) => s.id === workbook.activeSheetId) || workbook.sheets[0];
  }, [workbook.sheets, workbook.activeSheetId]);

  const [editingCell, setEditingCell] = useState<string | null>(null);
  const [cellInputDraft, setCellInputDraft] = useState('');
  const [renamingSheetId, setRenamingSheetId] = useState<string | null>(null);
  const [sheetNameDraft, setSheetNameDraft] = useState('');
  const [isMouseDown, setIsMouseDown] = useState(false);
  const [dragStartCell, setDragStartCell] = useState<string | null>(null);

  const gridContainerRef = useRef<HTMLDivElement>(null);
  const formulaInputRef = useRef<HTMLInputElement>(null);
  const cellInputRef = useRef<HTMLInputElement>(null);

  const t = getTranslation(lang);
  const isRTL = lang === 'ar';

  const rowCount = activeSheet?.rowCount || 45;
  const colCount = activeSheet?.colCount || 20;

  const activeCellData = activeSheet?.cells[workbook.selectedCell] || { raw: '' };

  // Calculate selected range cells
  const selectedCellList = useMemo(() => {
    if (!workbook.selectionRange) {
      return [workbook.selectedCell];
    }
    return expandCellRange(`${workbook.selectionRange.start}:${workbook.selectionRange.end}`);
  }, [workbook.selectedCell, workbook.selectionRange]);

  // Live Status Bar range stats (Average, Count, Sum, Min, Max)
  const rangeStats = useMemo(() => {
    if (!activeSheet) return { sum: 0, average: 0, count: 0, min: 0, max: 0, numericCount: 0 };
    return calculateRangeStats(selectedCellList, activeSheet.cells);
  }, [selectedCellList, activeSheet]);

  // Sync draft input when active cell changes
  useEffect(() => {
    if (editingCell !== workbook.selectedCell) {
      setCellInputDraft(activeCellData.raw || '');
    }
  }, [workbook.selectedCell, activeCellData.raw]);

  // Handle cell selection
  const handleCellClick = (cellRef: string, e: React.MouseEvent) => {
    if (editingCell && editingCell !== cellRef) {
      commitCellEdit();
    }
    if (e.shiftKey) {
      onSelectCell(workbook.selectedCell, {
        start: workbook.selectedCell,
        end: cellRef,
      });
    } else {
      onSelectCell(cellRef, null);
    }
  };

  const handleCellMouseDown = (cellRef: string) => {
    setIsMouseDown(true);
    setDragStartCell(cellRef);
    onSelectCell(cellRef, null);
  };

  const handleCellMouseEnter = (cellRef: string) => {
    if (isMouseDown && dragStartCell) {
      onSelectCell(dragStartCell, {
        start: dragStartCell,
        end: cellRef,
      });
    }
  };

  useEffect(() => {
    const handleMouseUp = () => {
      setIsMouseDown(false);
      setDragStartCell(null);
    };
    window.addEventListener('mouseup', handleMouseUp);
    return () => window.removeEventListener('mouseup', handleMouseUp);
  }, []);

  const handleCellDoubleClick = (cellRef: string) => {
    setEditingCell(cellRef);
    setCellInputDraft(activeSheet.cells[cellRef]?.raw || '');
    setTimeout(() => cellInputRef.current?.focus(), 50);
  };

  const commitCellEdit = () => {
    if (editingCell) {
      onUpdateCell(activeSheet.id, editingCell, { raw: cellInputDraft });
      setEditingCell(null);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent, cellRef: string) => {
    if (editingCell) {
      if (e.key === 'Enter') {
        e.preventDefault();
        commitCellEdit();
        // Move selection down
        const parsed = parseCellRef(cellRef);
        if (parsed && parsed.row < rowCount) {
          onSelectCell(`${parsed.col}${parsed.row + 1}`, null);
        }
      } else if (e.key === 'Tab') {
        e.preventDefault();
        commitCellEdit();
        // Move selection right
        const parsed = parseCellRef(cellRef);
        if (parsed && parsed.colIdx < colCount - 1) {
          onSelectCell(`${indexToColName(parsed.colIdx + 1)}${parsed.row}`, null);
        }
      } else if (e.key === 'Escape') {
        setEditingCell(null);
        setCellInputDraft(activeCellData.raw || '');
      }
      return;
    }

    // Navigation when not in edit mode
    const parsed = parseCellRef(cellRef);
    if (!parsed) return;

    if (e.key === 'ArrowDown' && parsed.row < rowCount) {
      e.preventDefault();
      onSelectCell(`${parsed.col}${parsed.row + 1}`, null);
    } else if (e.key === 'ArrowUp' && parsed.row > 1) {
      e.preventDefault();
      onSelectCell(`${parsed.col}${parsed.row - 1}`, null);
    } else if (e.key === 'ArrowRight' && parsed.colIdx < colCount - 1) {
      e.preventDefault();
      onSelectCell(`${indexToColName(parsed.colIdx + 1)}${parsed.row}`, null);
    } else if (e.key === 'ArrowLeft' && parsed.colIdx > 0) {
      e.preventDefault();
      onSelectCell(`${indexToColName(parsed.colIdx - 1)}${parsed.row}`, null);
    } else if (e.key === 'Enter' || e.key === 'F2') {
      e.preventDefault();
      handleCellDoubleClick(cellRef);
    } else if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      selectedCellList.forEach((ref) => {
        onUpdateCell(activeSheet.id, ref, { raw: '' });
      });
      setCellInputDraft('');
    } else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      // Start typing directly in cell
      setEditingCell(cellRef);
      setCellInputDraft(e.key);
      setTimeout(() => cellInputRef.current?.focus(), 20);
    }
  };

  // Formula bar commit
  const handleFormulaBarChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setCellInputDraft(e.target.value);
    onUpdateCell(activeSheet.id, workbook.selectedCell, { raw: e.target.value });
  };

  // Sheet renaming
  const handleStartRenameSheet = (sheet: ExcelSheet) => {
    setRenamingSheetId(sheet.id);
    setSheetNameDraft(sheet.name);
  };

  const handleFinishRenameSheet = () => {
    if (renamingSheetId && sheetNameDraft.trim()) {
      onRenameSheet(renamingSheetId, sheetNameDraft.trim());
      setRenamingSheetId(null);
    }
  };

  // Columns array
  const columns = useMemo(() => {
    return Array.from({ length: colCount }).map((_, idx) => indexToColName(idx));
  }, [colCount]);

  // Display columns (can reverse if RTL)
  const displayColumns = useMemo(() => {
    return isRTL ? [...columns].reverse() : columns;
  }, [columns, isRTL]);

  const rows = useMemo(() => {
    return Array.from({ length: rowCount }).map((_, idx) => idx + 1);
  }, [rowCount]);

  return (
    <div className="flex-1 flex flex-col h-full bg-[#f3f2f1] dark:bg-[#1b1a19] overflow-hidden select-none">
      {/* 1. Formula Bar */}
      <div className="bg-white dark:bg-[#252423] border-b border-gray-300 dark:border-[#3b3a39] px-2 py-1 flex items-center gap-2 text-xs shadow-sm shrink-0">
        {/* Active cell reference box */}
        <div className="font-mono font-semibold bg-gray-100 dark:bg-[#1b1a19] border border-gray-300 dark:border-[#3b3a39] rounded px-2.5 py-1 min-w-[54px] text-center text-gray-800 dark:text-gray-200">
          {workbook.selectedCell}
        </div>

        {/* fx button */}
        <button
          onClick={onOpenFormulaHelper}
          title={t.fxButton}
          className="font-serif italic font-bold text-[#107c41] dark:text-green-400 hover:bg-gray-100 dark:hover:bg-[#323130] rounded px-2 py-1 transition-colors cursor-pointer"
        >
          fx
        </button>

        {/* Live Formula Input Bar */}
        <div className="flex-1 flex items-center bg-gray-50 dark:bg-[#1b1a19] border border-gray-300 dark:border-[#3b3a39] rounded px-2 focus-within:ring-1 focus-within:ring-[#107c41] focus-within:border-[#107c41]">
          <input
            ref={formulaInputRef}
            type="text"
            value={cellInputDraft}
            onChange={handleFormulaBarChange}
            placeholder={t.formulaBarPlaceholder}
            className="w-full py-1 text-xs outline-none bg-transparent font-mono text-gray-800 dark:text-gray-100 select-text"
          />
        </div>
      </div>

      {/* 2. Spreadsheet Grid */}
      <div
        ref={gridContainerRef}
        tabIndex={0}
        className="flex-1 overflow-auto bg-white dark:bg-[#201f1e] relative outline-none"
        style={{
          transform: `scale(${workbook.zoom / 100})`,
          transformOrigin: 'top left',
          width: `${100 * (100 / workbook.zoom)}%`,
          height: `${100 * (100 / workbook.zoom)}%`,
        }}
      >
        <table className="border-collapse border-spacing-0 w-max table-fixed">
          {/* Header Row */}
          <thead>
            <tr className="bg-[#f3f2f1] dark:bg-[#292827] sticky top-0 z-20">
              {/* Top-left corner cell */}
              <th className="w-12 h-6 border-r border-b border-gray-300 dark:border-[#3b3a39] bg-[#e1dfdd] dark:bg-[#323130] sticky left-0 z-30 select-none">
                <div className="w-full h-full flex items-center justify-center text-[10px] text-gray-400">
                  ◢
                </div>
              </th>

              {/* Column headers (A, B, C...) */}
              {displayColumns.map((col) => {
                const isColActive = workbook.selectedCell.startsWith(col);
                const colWidth = activeSheet.colWidths?.[col] || 96;
                return (
                  <th
                    key={col}
                    style={{ width: `${colWidth}px`, minWidth: `${colWidth}px` }}
                    className={`h-6 border-r border-b border-gray-300 dark:border-[#3b3a39] text-[11px] font-medium transition-colors select-none ${
                      isColActive
                        ? 'bg-[#107c41] text-white font-bold'
                        : 'text-gray-600 dark:text-gray-300 hover:bg-gray-200 dark:hover:bg-[#3b3a39]'
                    }`}
                  >
                    {col}
                  </th>
                );
              })}
            </tr>
          </thead>

          {/* Grid Rows */}
          <tbody>
            {rows.map((row) => {
              const rowHeight = activeSheet.rowHeights?.[row] || 24;
              const isRowActive = workbook.selectedCell.endsWith(String(row));

              return (
                <tr key={row} style={{ height: `${rowHeight}px` }}>
                  {/* Row Number Header */}
                  <td
                    className={`w-12 border-r border-b border-gray-300 dark:border-[#3b3a39] text-[11px] text-center sticky left-0 z-10 transition-colors select-none ${
                      isRowActive
                        ? 'bg-[#107c41] text-white font-bold'
                        : 'bg-[#f3f2f1] dark:bg-[#292827] text-gray-600 dark:text-gray-400'
                    }`}
                  >
                    {row}
                  </td>

                  {/* Grid Cells in this row */}
                  {displayColumns.map((col) => {
                    const cellRef = `${col}${row}`;
                    const cellData = activeSheet.cells[cellRef] || { raw: '' };
                    const isSelected = workbook.selectedCell === cellRef;
                    const inRange = selectedCellList.includes(cellRef);
                    const isEditing = editingCell === cellRef;

                    // Computed value
                    const computedValue = evaluateCell(cellRef, activeSheet.cells);
                    const formattedDisplay = formatCellValue(
                      computedValue,
                      cellData.format,
                      cellData.decimals,
                      lang
                    );

                    // Cell styles
                    const cellStyle: React.CSSProperties = {
                      fontWeight: cellData.bold ? 'bold' : 'normal',
                      fontStyle: cellData.italic ? 'italic' : 'normal',
                      textDecoration: cellData.underline ? 'underline' : 'none',
                      color: cellData.color || undefined,
                      backgroundColor: cellData.bgColor || undefined,
                      textAlign: cellData.align || (typeof computedValue === 'number' ? 'right' : 'left'),
                      fontFamily: cellData.fontFamily || undefined,
                      fontSize: cellData.fontSize ? `${cellData.fontSize}pt` : undefined,
                    };

                    return (
                      <td
                        key={cellRef}
                        onMouseDown={() => handleCellMouseDown(cellRef)}
                        onMouseEnter={() => handleCellMouseEnter(cellRef)}
                        onClick={(e) => handleCellClick(cellRef, e)}
                        onDoubleClick={() => handleCellDoubleClick(cellRef)}
                        onKeyDown={(e) => handleKeyDown(e, cellRef)}
                        tabIndex={0}
                        style={cellStyle}
                        className={`px-2 text-xs border-r border-b relative select-none outline-none transition-colors truncate ${
                          workbook.showGridlines
                            ? 'border-gray-200 dark:border-[#333230]'
                            : 'border-transparent'
                        } ${
                          isSelected
                            ? 'ring-2 ring-[#107c41] ring-inset z-10 bg-green-50/50 dark:bg-green-950/20'
                            : inRange
                            ? 'bg-green-100/40 dark:bg-green-900/20'
                            : ''
                        }`}
                      >
                        {isEditing ? (
                          <input
                            ref={cellInputRef}
                            type="text"
                            value={cellInputDraft}
                            onChange={(e) => setCellInputDraft(e.target.value)}
                            onBlur={commitCellEdit}
                            onKeyDown={(e) => handleKeyDown(e, cellRef)}
                            autoFocus
                            className="absolute inset-0 w-full h-full px-2 text-xs bg-white dark:bg-[#1b1a19] text-gray-900 dark:text-gray-100 outline-none ring-2 ring-[#107c41] select-text z-20 font-mono"
                          />
                        ) : (
                          <span>{formattedDisplay}</span>
                        )}

                        {/* Excel bottom-right handle on selected cell */}
                        {isSelected && !isEditing && (
                          <div className="absolute -bottom-1 -right-1 w-2 h-2 bg-[#107c41] border border-white z-20 cursor-crosshair pointer-events-none" />
                        )}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* 3. Sheet Tabs Rail */}
      <div className="bg-[#f3f2f1] dark:bg-[#252423] border-t border-gray-300 dark:border-[#3b3a39] px-2 py-1 flex items-center justify-between text-xs select-none shrink-0">
        <div className="flex items-center gap-1 overflow-x-auto">
          {workbook.sheets.map((sheet) => {
            const isActive = sheet.id === workbook.activeSheetId;
            const isRenaming = renamingSheetId === sheet.id;

            return (
              <div
                key={sheet.id}
                onDoubleClick={() => handleStartRenameSheet(sheet)}
                onClick={() => onSetActiveSheet(sheet.id)}
                className={`group px-3 py-1 rounded-t flex items-center gap-2 cursor-pointer transition-colors border-t border-x ${
                  isActive
                    ? 'bg-white dark:bg-[#1b1a19] border-gray-300 dark:border-[#3b3a39] font-semibold text-[#107c41] dark:text-green-400 shadow-xs'
                    : 'bg-transparent border-transparent text-gray-600 dark:text-gray-400 hover:bg-gray-200 dark:hover:bg-[#323130]'
                }`}
              >
                {isRenaming ? (
                  <input
                    type="text"
                    value={sheetNameDraft}
                    onChange={(e) => setSheetNameDraft(e.target.value)}
                    onBlur={handleFinishRenameSheet}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') handleFinishRenameSheet();
                      if (e.key === 'Escape') setRenamingSheetId(null);
                    }}
                    autoFocus
                    className="w-20 px-1 py-0.5 text-xs rounded border border-[#107c41] outline-none bg-white dark:bg-[#252423]"
                  />
                ) : (
                  <span>{sheet.name}</span>
                )}

                {/* Delete button (if more than 1 sheet) */}
                {workbook.sheets.length > 1 && (
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      onDeleteSheet(sheet.id);
                    }}
                    title={t.deleteSheet}
                    className="opacity-0 group-hover:opacity-100 hover:text-red-500 transition-opacity"
                  >
                    <X className="w-3 h-3" />
                  </button>
                )}
              </div>
            );
          })}

          {/* Add Sheet (+) Button */}
          <button
            onClick={onAddSheet}
            title={t.addSheet}
            className="p-1 rounded-full hover:bg-gray-200 dark:hover:bg-[#323130] text-gray-600 dark:text-gray-300 transition-colors ml-1 cursor-pointer"
          >
            <Plus className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* 4. Live Excel Status Bar */}
      <div className="status-bar no-print bg-[#e1dfdd] dark:bg-[#201f1e] border-t border-gray-300 dark:border-[#3b3a39] px-4 py-1 flex items-center justify-between text-xs text-gray-700 dark:text-gray-300 select-none shrink-0">
        {/* Statistics of selected cells */}
        <div className="flex items-center gap-4">
          <span className="text-[11px] font-mono text-gray-500">READY</span>
          {rangeStats.numericCount > 0 && (
            <>
              <div className="flex items-center gap-1">
                <span className="text-gray-500">{t.average}:</span>
                <span className="font-semibold font-mono">
                  {rangeStats.average.toLocaleString()}
                </span>
              </div>
              <div className="flex items-center gap-1">
                <span className="text-gray-500">{t.count}:</span>
                <span className="font-semibold font-mono">{rangeStats.count}</span>
              </div>
              <div className="flex items-center gap-1">
                <span className="text-gray-500">{t.sum}:</span>
                <span className="font-semibold font-mono text-[#107c41] dark:text-green-400">
                  {rangeStats.sum.toLocaleString()}
                </span>
              </div>
            </>
          )}
        </div>

        {/* Zoom Slider */}
        <div className="flex items-center gap-2">
          <button
            onClick={() => onSetZoom(Math.max(50, workbook.zoom - 10))}
            className="p-0.5 rounded hover:bg-gray-200 dark:hover:bg-[#323130]"
            title="Zoom Out"
          >
            <ZoomOut className="w-3.5 h-3.5" />
          </button>

          <input
            type="range"
            min="50"
            max="200"
            step="5"
            value={workbook.zoom}
            onChange={(e) => onSetZoom(parseInt(e.target.value, 10))}
            className="w-24 h-1.5 bg-gray-300 dark:bg-gray-700 rounded-lg appearance-none cursor-pointer accent-[#107c41]"
          />

          <button
            onClick={() => onSetZoom(Math.min(200, workbook.zoom + 10))}
            className="p-0.5 rounded hover:bg-gray-200 dark:hover:bg-[#323130]"
            title="Zoom In"
          >
            <ZoomIn className="w-3.5 h-3.5" />
          </button>

          <span className="font-mono text-[11px] w-9 text-right">{workbook.zoom}%</span>
        </div>
      </div>
    </div>
  );
};

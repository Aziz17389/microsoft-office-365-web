export type Language = 'en' | 'ar';
export type Theme = 'light' | 'dark';
export type AppMode = 'launcher' | 'word' | 'excel';
export type RibbonTab = 'file' | 'home' | 'insert' | 'layout' | 'view';

export type PlanTier = 'free' | 'pro';

export type GatedFeature =
  | 'unlimited_time'
  | 'tables'
  | 'save_export'
  | 'font_family'
  | 'text_size'
  | 'page_size'
  | 'hyperlinks';

export type WordPageMargins = 'normal' | 'narrow' | 'wide';
export type WordPageOrientation = 'portrait' | 'landscape';
export type WordPaperSize = 'A4' | 'Letter' | 'Legal';

export interface WordEditorState {
  id: string;
  title: string;
  htmlContent: string;
  margins: WordPageMargins;
  orientation: WordPageOrientation;
  paperSize: WordPaperSize;
  zoom: number;
  showRuler: boolean;
  showGridlines: boolean;
  isZenMode: boolean;
}

export type CellFormatType = 'general' | 'number' | 'currency' | 'percent';

export interface CellData {
  raw: string;
  computed?: string | number;
  format?: CellFormatType;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  bgColor?: string;
  align?: 'left' | 'center' | 'right';
  decimals?: number;
  fontFamily?: string;
  fontSize?: number;
}

export interface ExcelSheet {
  id: string;
  name: string;
  cells: Record<string, CellData>;
  colWidths: Record<string, number>;
  rowHeights: Record<number, number>;
  rowCount: number;
  colCount: number;
}

export interface ExcelWorkbookState {
  id: string;
  title: string;
  sheets: ExcelSheet[];
  activeSheetId: string;
  selectedCell: string;
  selectionRange: {
    start: string;
    end: string;
  } | null;
  zoom: number;
  showGridlines: boolean;
  isZenMode: boolean;
}

export interface RecentDocument {
  id: string;
  title: string;
  app: 'word' | 'excel';
  updatedAt: number;
  preview: string;
  size: string;
}

export interface FormattingState {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  subscript: boolean;
  superscript: boolean;
  fontName: string;
  fontSize: string;
  foreColor: string;
  hiliteColor: string;
  align: 'left' | 'center' | 'right' | 'justify';
}

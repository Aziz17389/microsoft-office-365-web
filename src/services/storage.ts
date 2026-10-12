import { ExcelWorkbookState, RecentDocument, WordEditorState } from '../types/office';

const RECENT_DOCS_KEY = 'office365_recent_documents';
const WORD_PREFIX = 'office365_word_doc_';
const EXCEL_PREFIX = 'office365_excel_doc_';

// Initial Word templates
export const WORD_TEMPLATES: Record<string, { title: string; html: string }> = {
  blank: {
    title: 'Document 1',
    html: '<p>Type your text here...</p>',
  },
  proposal: {
    title: 'Project Proposal',
    html: `
      <h1 style="color: #185abd; font-size: 28px; font-weight: bold; margin-bottom: 8px;">Project Phoenix: Enterprise Modernization</h1>
      <p style="color: #666; font-size: 14px; margin-bottom: 24px;">Prepared by Strategy Team | Date: October 2026</p>
      
      <h2 style="font-size: 20px; font-weight: 600; color: #333; margin-top: 16px; margin-bottom: 8px;">1. Executive Summary</h2>
      <p style="line-height: 1.6; margin-bottom: 16px;">This proposal outlines the strategic roadmap for upgrading the company digital productivity infrastructure. Implementing web-first Office suites will improve cross-functional collaboration and reduce operational friction by 35%.</p>
      
      <h2 style="font-size: 20px; font-weight: 600; color: #333; margin-top: 16px; margin-bottom: 8px;">2. Project Objectives</h2>
      <ul style="margin-left: 20px; margin-bottom: 16px; line-height: 1.6;">
        <li>Migrate legacy documentation to unified browser workspace</li>
        <li>Enable seamless Arabic/English dual language editing</li>
        <li>Achieve 99.9% uptime with instant local storage resilience</li>
      </ul>

      <h2 style="font-size: 20px; font-weight: 600; color: #333; margin-top: 16px; margin-bottom: 8px;">3. Project Milestones & Budget</h2>
      <table style="width: 100%; border-collapse: collapse; margin-top: 12px; margin-bottom: 20px;">
        <thead>
          <tr style="background-color: #f3f2f1; text-align: left;">
            <th style="border: 1px solid #d1d5db; padding: 10px;">Phase</th>
            <th style="border: 1px solid #d1d5db; padding: 10px;">Deliverable</th>
            <th style="border: 1px solid #d1d5db; padding: 10px;">Target Date</th>
            <th style="border: 1px solid #d1d5db; padding: 10px;">Est. Cost</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="border: 1px solid #d1d5db; padding: 8px;">Phase 1</td>
            <td style="border: 1px solid #d1d5db; padding: 8px;">Architecture & Design</td>
            <td style="border: 1px solid #d1d5db; padding: 8px;">Nov 2026</td>
            <td style="border: 1px solid #d1d5db; padding: 8px;">$15,000</td>
          </tr>
          <tr>
            <td style="border: 1px solid #d1d5db; padding: 8px;">Phase 2</td>
            <td style="border: 1px solid #d1d5db; padding: 8px;">Editor Core & Spreadsheets</td>
            <td style="border: 1px solid #d1d5db; padding: 8px;">Jan 2027</td>
            <td style="border: 1px solid #d1d5db; padding: 8px;">$35,000</td>
          </tr>
          <tr>
            <td style="border: 1px solid #d1d5db; padding: 8px;">Phase 3</td>
            <td style="border: 1px solid #d1d5db; padding: 8px;">Deployment & Training</td>
            <td style="border: 1px solid #d1d5db; padding: 8px;">Mar 2027</td>
            <td style="border: 1px solid #d1d5db; padding: 8px;">$10,000</td>
          </tr>
        </tbody>
      </table>
    `,
  },
  meetingNotes: {
    title: 'Weekly Executive Briefing',
    html: `
      <h1 style="color: #185abd; font-size: 26px; font-weight: bold; margin-bottom: 4px;">Executive Team Weekly Meeting</h1>
      <p style="color: #777; margin-bottom: 20px;">October 12, 2026 | Conference Room A & Teams</p>
      
      <h3 style="font-size: 18px; font-weight: 600; color: #222;">Attendees</h3>
      <p style="margin-bottom: 16px;">Sara Al-Mansoor (CEO), Tariq Vance (VP Product), David Miller (Lead Architect), Fatima Noor (Engineering Lead)</p>

      <h3 style="font-size: 18px; font-weight: 600; color: #222;">Agenda Topics</h3>
      <ol style="margin-left: 20px; margin-bottom: 20px; line-height: 1.6;">
        <li>Q3 Performance review & customer adoption stats</li>
        <li>Review of the new Excel recalculation engine</li>
        <li>Bilingual RTL Arabic UX evaluation</li>
      </ol>

      <h3 style="font-size: 18px; font-weight: 600; color: #222;">Key Action Items</h3>
      <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">
        <thead>
          <tr style="background-color: #f3f2f1; text-align: left;">
            <th style="border: 1px solid #ccc; padding: 8px;">Action Item</th>
            <th style="border: 1px solid #ccc; padding: 8px;">Assignee</th>
            <th style="border: 1px solid #ccc; padding: 8px;">Due Date</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="border: 1px solid #ccc; padding: 8px;">Finalize formula builder dialog</td>
            <td style="border: 1px solid #ccc; padding: 8px;">Fatima Noor</td>
            <td style="border: 1px solid #ccc; padding: 8px;">Friday</td>
          </tr>
          <tr>
            <td style="border: 1px solid #ccc; padding: 8px;">Prepare QA test matrix for RTL grid</td>
            <td style="border: 1px solid #ccc; padding: 8px;">David Miller</td>
            <td style="border: 1px solid #ccc; padding: 8px;">Oct 20</td>
          </tr>
        </tbody>
      </table>
    `,
  },
  resume: {
    title: 'Professional Resume',
    html: `
      <div style="text-align: center; margin-bottom: 24px; border-bottom: 2px solid #185abd; padding-bottom: 16px;">
        <h1 style="font-size: 32px; font-weight: bold; color: #185abd; margin: 0;">ALEXANDER WRIGHT</h1>
        <p style="font-size: 16px; color: #555; margin: 4px 0;">Senior Full-Stack Engineer & Cloud Solutions Architect</p>
        <p style="font-size: 13px; color: #777;">alex.wright@example.com | +1 (555) 234-5678 | San Francisco, CA</p>
      </div>

      <h2 style="font-size: 18px; color: #185abd; border-bottom: 1px solid #ddd; padding-bottom: 4px;">PROFESSIONAL SUMMARY</h2>
      <p style="line-height: 1.6; margin-bottom: 20px;">Dedicated software engineer with 8+ years of experience engineering high-scale web platforms, desktop web simulators, and real-time computation engines. Proven track record of architecting mission-critical SaaS applications.</p>

      <h2 style="font-size: 18px; color: #185abd; border-bottom: 1px solid #ddd; padding-bottom: 4px;">EXPERIENCE</h2>
      <div style="margin-bottom: 16px;">
        <h3 style="font-size: 16px; font-weight: 600; margin: 0;">Lead Web Engineer — CloudMatrix Inc.</h3>
        <p style="color: #666; font-size: 13px; margin: 2px 0 8px 0;">2022 – Present | San Francisco, CA</p>
        <ul style="margin-left: 20px; line-height: 1.5;">
          <li>Architected browser-based document editor used by over 250,000 active monthly professionals.</li>
          <li>Engineered dynamic formula evaluation engine with sub-10ms recalculation speed.</li>
        </ul>
      </div>

      <h2 style="font-size: 18px; color: #185abd; border-bottom: 1px solid #ddd; padding-bottom: 4px;">EDUCATION</h2>
      <p style="margin: 4px 0;"><strong>B.S. in Computer Science</strong> — Stanford University (2018)</p>
    `,
  },
};

// Initial Excel templates
export const EXCEL_TEMPLATES: Record<string, { title: string; cells: Record<string, any> }> = {
  blank: {
    title: 'Book 1',
    cells: {},
  },
  budget: {
    title: 'Monthly Budget 2026',
    cells: {
      A1: { raw: 'Monthly Household Budget', bold: true, fontSize: 16, color: '#107c41' },
      A3: { raw: 'Category', bold: true, bgColor: '#e1dfdd' },
      B3: { raw: 'Budgeted', bold: true, bgColor: '#e1dfdd', align: 'right' },
      C3: { raw: 'Actual', bold: true, bgColor: '#e1dfdd', align: 'right' },
      D3: { raw: 'Difference', bold: true, bgColor: '#e1dfdd', align: 'right' },
      // Income
      A4: { raw: 'Salary & Wages' },
      B4: { raw: '5000', format: 'currency', align: 'right' },
      C4: { raw: '5200', format: 'currency', align: 'right' },
      D4: { raw: '=C4-B4', format: 'currency', align: 'right' },
      // Expenses
      A5: { raw: 'Housing / Rent' },
      B5: { raw: '1800', format: 'currency', align: 'right' },
      C5: { raw: '1800', format: 'currency', align: 'right' },
      D5: { raw: '=B5-C5', format: 'currency', align: 'right' },
      A6: { raw: 'Groceries' },
      B6: { raw: '600', format: 'currency', align: 'right' },
      C6: { raw: '650', format: 'currency', align: 'right' },
      D6: { raw: '=B6-C6', format: 'currency', align: 'right' },
      A7: { raw: 'Utilities & Internet' },
      B7: { raw: '300', format: 'currency', align: 'right' },
      C7: { raw: '280', format: 'currency', align: 'right' },
      D7: { raw: '=B7-C7', format: 'currency', align: 'right' },
      A8: { raw: 'Transportation' },
      B8: { raw: '400', format: 'currency', align: 'right' },
      C8: { raw: '350', format: 'currency', align: 'right' },
      D8: { raw: '=B8-C8', format: 'currency', align: 'right' },
      // Totals
      A10: { raw: 'Total Expenses', bold: true },
      B10: { raw: '=SUM(B5:B8)', bold: true, format: 'currency', align: 'right' },
      C10: { raw: '=SUM(C5:C8)', bold: true, format: 'currency', align: 'right' },
      D10: { raw: '=SUM(D5:D8)', bold: true, format: 'currency', align: 'right' },
      A12: { raw: 'Net Monthly Savings', bold: true, color: '#107c41' },
      C12: { raw: '=C4-C10', bold: true, format: 'currency', align: 'right', color: '#107c41' },
    },
  },
  sales: {
    title: 'Sales & Inventory Tracker',
    cells: {
      A1: { raw: 'Q4 Product Sales Report', bold: true, fontSize: 16, color: '#107c41' },
      A3: { raw: 'Product Name', bold: true, bgColor: '#dff6dd' },
      B3: { raw: 'Units Sold', bold: true, bgColor: '#dff6dd', align: 'right' },
      C3: { raw: 'Unit Price', bold: true, bgColor: '#dff6dd', align: 'right' },
      D3: { raw: 'Total Revenue', bold: true, bgColor: '#dff6dd', align: 'right' },
      // Rows
      A4: { raw: 'Surface Pro Laptop' },
      B4: { raw: '120', align: 'right' },
      C4: { raw: '999', format: 'currency', align: 'right' },
      D4: { raw: '=B4*C4', format: 'currency', align: 'right', bold: true },
      A5: { raw: 'Wireless Mechanical Keyboard' },
      B5: { raw: '340', align: 'right' },
      C5: { raw: '129', format: 'currency', align: 'right' },
      D5: { raw: '=B5*C5', format: 'currency', align: 'right', bold: true },
      A6: { raw: 'Ergonomic Precision Mouse' },
      B6: { raw: '450', align: 'right' },
      C6: { raw: '79', format: 'currency', align: 'right' },
      D6: { raw: '=B6*C6', format: 'currency', align: 'right', bold: true },
      A7: { raw: '4K UltraHD Monitor 32"' },
      B7: { raw: '85', align: 'right' },
      C7: { raw: '499', format: 'currency', align: 'right' },
      D7: { raw: '=B7*C7', format: 'currency', align: 'right', bold: true },
      // Summary
      A9: { raw: 'Total Units Sold', bold: true },
      B9: { raw: '=SUM(B4:B7)', bold: true, align: 'right' },
      A10: { raw: 'Gross Total Revenue', bold: true, color: '#107c41' },
      D10: { raw: '=SUM(D4:D7)', bold: true, format: 'currency', align: 'right', color: '#107c41' },
      A11: { raw: 'Average Item Price', bold: true },
      C11: { raw: '=AVERAGE(C4:C7)', bold: true, format: 'currency', align: 'right' },
    },
  },
  gradebook: {
    title: 'Student Gradebook',
    cells: {
      A1: { raw: 'Computer Science 101 - Grades', bold: true, fontSize: 16, color: '#107c41' },
      A3: { raw: 'Student Name', bold: true, bgColor: '#f3f2f1' },
      B3: { raw: 'Quiz 1', bold: true, bgColor: '#f3f2f1', align: 'right' },
      C3: { raw: 'Midterm', bold: true, bgColor: '#f3f2f1', align: 'right' },
      D3: { raw: 'Final Exam', bold: true, bgColor: '#f3f2f1', align: 'right' },
      E3: { raw: 'Average Score', bold: true, bgColor: '#f3f2f1', align: 'right' },
      F3: { raw: 'Status', bold: true, bgColor: '#f3f2f1', align: 'center' },
      // Students
      A4: { raw: 'Youssef Ahmed' },
      B4: { raw: '92', align: 'right' },
      C4: { raw: '88', align: 'right' },
      D4: { raw: '95', align: 'right' },
      E4: { raw: '=AVERAGE(B4:D4)', bold: true, align: 'right' },
      F4: { raw: '=IF(E4>=60, "Pass", "Fail")', bold: true, align: 'center', color: '#107c41' },
      A5: { raw: 'Laila Al-Khatib' },
      B5: { raw: '85', align: 'right' },
      C5: { raw: '90', align: 'right' },
      D5: { raw: '92', align: 'right' },
      E5: { raw: '=AVERAGE(B5:D5)', bold: true, align: 'right' },
      F5: { raw: '=IF(E5>=60, "Pass", "Fail")', bold: true, align: 'center', color: '#107c41' },
      A6: { raw: 'Marcus Chen' },
      B6: { raw: '78', align: 'right' },
      C6: { raw: '82', align: 'right' },
      D6: { raw: '80', align: 'right' },
      E6: { raw: '=AVERAGE(B6:D6)', bold: true, align: 'right' },
      F6: { raw: '=IF(E6>=60, "Pass", "Fail")', bold: true, align: 'center', color: '#107c41' },
      // Class summary
      A8: { raw: 'Class Average', bold: true },
      E8: { raw: '=AVERAGE(E4:E6)', bold: true, align: 'right', color: '#107c41' },
      A9: { raw: 'Highest Grade', bold: true },
      E9: { raw: '=MAX(E4:E6)', bold: true, align: 'right' },
      A10: { raw: 'Lowest Grade', bold: true },
      E10: { raw: '=MIN(E4:E6)', bold: true, align: 'right' },
    },
  },
};

// Storage APIs
export function getRecentDocuments(): RecentDocument[] {
  try {
    const raw = localStorage.getItem(RECENT_DOCS_KEY);
    if (!raw) {
      // Seed with sample initial recent docs if empty
      const initial: RecentDocument[] = [
        {
          id: 'sample-word-proposal',
          title: 'Project Proposal',
          app: 'word',
          updatedAt: Date.now() - 1000 * 60 * 35,
          preview: 'Executive roadmap for enterprise modern productivity...',
          size: '14.2 KB',
        },
        {
          id: 'sample-excel-budget',
          title: 'Monthly Budget 2026',
          app: 'excel',
          updatedAt: Date.now() - 1000 * 60 * 120,
          preview: 'Income, Expenses, Net Savings summary calculation...',
          size: '8.6 KB',
        },
      ];
      localStorage.setItem(RECENT_DOCS_KEY, JSON.stringify(initial));
      // Save their contents
      saveWordDocument({
        id: 'sample-word-proposal',
        title: WORD_TEMPLATES.proposal.title,
        htmlContent: WORD_TEMPLATES.proposal.html,
        margins: 'normal',
        orientation: 'portrait',
        paperSize: 'A4',
        zoom: 100,
        showRuler: true,
        showGridlines: false,
        isZenMode: false,
      });
      saveExcelWorkbook({
        id: 'sample-excel-budget',
        title: EXCEL_TEMPLATES.budget.title,
        sheets: [
          {
            id: 'sheet-1',
            name: 'Budget',
            cells: EXCEL_TEMPLATES.budget.cells,
            colWidths: { A: 200, B: 120, C: 120, D: 120 },
            rowHeights: {},
            rowCount: 40,
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
      return initial;
    }
    return JSON.parse(raw);
  } catch {
    return [];
  }
}

export function saveRecentDocumentMeta(doc: RecentDocument) {
  try {
    const current = getRecentDocuments().filter((d) => d.id !== doc.id);
    const updated = [doc, ...current].slice(0, 20);
    localStorage.setItem(RECENT_DOCS_KEY, JSON.stringify(updated));
  } catch (err) {
    console.error('Error saving recent doc meta:', err);
  }
}

export function deleteRecentDocument(id: string) {
  try {
    const current = getRecentDocuments().filter((d) => d.id !== id);
    localStorage.setItem(RECENT_DOCS_KEY, JSON.stringify(current));
    localStorage.removeItem(WORD_PREFIX + id);
    localStorage.removeItem(EXCEL_PREFIX + id);
  } catch (err) {
    console.error('Error deleting document:', err);
  }
}

export function saveWordDocument(doc: WordEditorState) {
  try {
    localStorage.setItem(WORD_PREFIX + doc.id, JSON.stringify(doc));
    // Calculate approximate size
    const sizeBytes = new Blob([JSON.stringify(doc)]).size;
    const sizeStr = sizeBytes > 1024 ? `${(sizeBytes / 1024).toFixed(1)} KB` : `${sizeBytes} B`;

    // Strip tags for preview
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = doc.htmlContent;
    const preview = (tempDiv.textContent || tempDiv.innerText || '').slice(0, 90) + '...';

    saveRecentDocumentMeta({
      id: doc.id,
      title: doc.title,
      app: 'word',
      updatedAt: Date.now(),
      preview: preview.trim() || 'Empty document',
      size: sizeStr,
    });
  } catch (err) {
    console.error('Failed to save Word document:', err);
  }
}

export function loadWordDocument(id: string): WordEditorState | null {
  try {
    const raw = localStorage.getItem(WORD_PREFIX + id);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

export function saveExcelWorkbook(workbook: ExcelWorkbookState) {
  try {
    localStorage.setItem(EXCEL_PREFIX + workbook.id, JSON.stringify(workbook));
    const sizeBytes = new Blob([JSON.stringify(workbook)]).size;
    const sizeStr = sizeBytes > 1024 ? `${(sizeBytes / 1024).toFixed(1)} KB` : `${sizeBytes} B`;

    const activeSheet = workbook.sheets.find((s) => s.id === workbook.activeSheetId) || workbook.sheets[0];
    const preview = `${workbook.sheets.length} Sheet(s) • ${Object.keys(activeSheet?.cells || {}).length} Cells populated`;

    saveRecentDocumentMeta({
      id: workbook.id,
      title: workbook.title,
      app: 'excel',
      updatedAt: Date.now(),
      preview,
      size: sizeStr,
    });
  } catch (err) {
    console.error('Failed to save Excel workbook:', err);
  }
}

export function loadExcelWorkbook(id: string): ExcelWorkbookState | null {
  try {
    const raw = localStorage.getItem(EXCEL_PREFIX + id);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

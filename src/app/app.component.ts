import { Component, ViewChild, AfterViewInit, ElementRef, NgZone } from '@angular/core';
import { SelectionModel } from '@angular/cdk/collections';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { MatTableModule, MatTableDataSource } from '@angular/material/table';
import { MatCheckboxModule } from '@angular/material/checkbox';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { MatToolbarModule } from '@angular/material/toolbar';
import { MatCardModule } from '@angular/material/card';
import { MatPaginatorModule, MatPaginator } from '@angular/material/paginator';
import { MatSortModule, MatSort } from '@angular/material/sort';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { MatDatepickerModule } from '@angular/material/datepicker';
import { MatNativeDateModule } from '@angular/material/core';
import { MatSelectModule } from '@angular/material/select';
import { MatChipsModule } from '@angular/material/chips';
import { MatExpansionModule } from '@angular/material/expansion';
import * as XLSX from 'xlsx';

export interface ApplicantRecord {
  applicant: string;
  acronym: string;
  entityType: string;
  country: string;
  address: string;
  nolStatus: string;
  hatyjaReviewComments: string;
  redFlags: string;
  passed: string;
  // Middle columns
  submittedAt?: Date | string;
  preScreening?: string;
  profiles?: string;
  // End columns
  djResult?: string;
  djReportNumber?: string;
  djReportLink?: string;
  djTruePositive?: string;
  djFalsePositive?: string;
  escalationRequired?: string;
  hatyjaComments?: string;
  // UI flags
  isEditingReview?: boolean;
  isEditingHatyjaComments?: boolean;
}

export interface DashboardStats {
  total: number;
  countries: number;
  accepted: number;
  rejected: number;
  pending: number;
  redFlags: number;
  acceptedPct: number;
  rejectedPct: number;
  pendingPct: number;
  redFlagsPct: number;
  topCountriesRedFlags: { name: string; count: number }[];
}

const ELEMENT_DATA: ApplicantRecord[] = [
  { applicant: 'Acme Corp', acronym: 'AC', entityType: 'Corporation', submittedAt: new Date('2023-01-15'), preScreening: 'Pass', profiles: 'https://example.com/acme', country: 'USA', address: '123 Main St', nolStatus: 'Active', hatyjaReviewComments: 'Looks good', redFlags: 'None', passed: 'Accepted', djResult: 'Clean', djReportNumber: 'DJ-101', djReportLink: 'https://dj.com/101', djTruePositive: 'No', djFalsePositive: 'No', escalationRequired: 'No', hatyjaComments: 'Ready for full review' },
  { applicant: 'Global Tech', acronym: 'GT', entityType: 'LLC', submittedAt: new Date('2023-02-10'), preScreening: 'Pending', profiles: 'https://example.com/gt', country: 'UK', address: '', nolStatus: 'Pending', hatyjaReviewComments: 'Needs more info', redFlags: 'Incomplete documents', passed: '', djResult: 'Warning', djReportNumber: 'DJ-202', djReportLink: 'https://dj.com/202', djTruePositive: 'Maybe', djFalsePositive: 'No', escalationRequired: 'Yes', hatyjaComments: 'Contact client' },
  { applicant: 'HealthPlus', acronym: 'HP', entityType: 'Non-Profit', submittedAt: new Date('2023-03-05'), preScreening: 'Fail', profiles: 'https://example.com/hp', country: 'Canada', address: '456 Maple Ave', nolStatus: 'Inactive', hatyjaReviewComments: 'Check compliance', redFlags: 'Expired license', passed: 'Rejected', djResult: 'Critical', djReportNumber: 'DJ-303', djReportLink: 'https://dj.com/303', djTruePositive: 'Yes', djFalsePositive: 'No', escalationRequired: 'Immediate', hatyjaComments: 'Review blocked' },
  { applicant: 'FinServe', acronym: 'FS', entityType: 'Partnership', submittedAt: new Date('2023-04-20'), preScreening: 'Pass', profiles: 'https://example.com/fs', country: 'Australia', address: '789 Wall St', nolStatus: 'Active', hatyjaReviewComments: 'All clear', redFlags: 'None', passed: 'Accepted', djResult: 'Clean', djReportNumber: 'DJ-404', djReportLink: 'https://dj.com/404', djTruePositive: 'No', djFalsePositive: 'No', escalationRequired: 'No', hatyjaComments: 'Good to go' },
  { applicant: 'ConstructCo', acronym: 'CC', entityType: 'Corporation', submittedAt: new Date('2023-05-12'), preScreening: 'Review', profiles: 'https://example.com/cc', country: 'Germany', address: '', nolStatus: 'Pending', hatyjaReviewComments: 'Pending review', redFlags: 'High risk', passed: '', djResult: 'Manual', djReportNumber: 'DJ-505', djReportLink: 'https://dj.com/505', djTruePositive: 'No', djFalsePositive: 'Yes', escalationRequired: 'To DRMC', hatyjaComments: 'Check financial year 2022' }
];

const STORAGE_KEY = 'gfc_applicant_data';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatTableModule,
    MatCheckboxModule,
    MatButtonModule,
    MatIconModule,
    MatToolbarModule,
    MatCardModule,
    MatPaginatorModule,
    MatFormFieldModule,
    MatInputModule,
    MatDatepickerModule,
    MatNativeDateModule,
    MatSelectModule,
    MatChipsModule,
    MatExpansionModule
  ],
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements AfterViewInit {
  title = 'angular-excel-app';
  displayedColumns: string[] = [];
  dataSource = new MatTableDataSource<ApplicantRecord>(ELEMENT_DATA);

  @ViewChild(MatPaginator) paginator!: MatPaginator;
  @ViewChild(MatSort) sort!: MatSort;

  editingReviewElement: ApplicantRecord | null = null;
  selection = new SelectionModel<ApplicantRecord>(true, []);
  selectedColumnKeys: string[] = [];

  get stats(): DashboardStats {
    const data = this.dataSource.data;
    const total = data.length;
    if (total === 0) {
      return {
        total: 0,
        countries: 0,
        accepted: 0,
        rejected: 0,
        pending: 0,
        redFlags: 0,
        acceptedPct: 0,
        rejectedPct: 0,
        pendingPct: 0,
        redFlagsPct: 0,
        topCountriesRedFlags: []
      };
    }

    const countries = new Set(data.map(d => d.country).filter(c => !!c)).size;
    const accepted = data.filter(d => d.passed === 'Accepted').length;
    const rejected = data.filter(d => d.passed === 'Rejected').length;
    const pending = total - accepted - rejected;
    const redFlags = data.filter(d => d.redFlags && d.redFlags !== 'None').length;

    return {
      total,
      countries,
      accepted,
      rejected,
      pending,
      redFlags,
      acceptedPct: (accepted / total) * 100,
      rejectedPct: (rejected / total) * 100,
      pendingPct: (pending / total) * 100,
      redFlagsPct: (redFlags / total) * 100,
      topCountriesRedFlags: this.getTopCountriesRedFlags(data)
    };
  }

  private getTopCountriesRedFlags(data: ApplicantRecord[]) {
    const counts: Record<string, number> = {};
    data.filter(d => d.redFlags && d.redFlags !== 'None' && d.country)
      .forEach(d => {
        counts[d.country] = (counts[d.country] || 0) + 1;
      });

    return Object.keys(counts)
      .map(name => ({ name, count: counts[name] }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 3);
  }

  constructor(private ngZone: NgZone) {
    this.syncSelectedKeys();
    this.updateDisplayedColumns();
  }

  ngAfterViewInit() {
    this.dataSource.paginator = this.paginator;
    this.dataSource.sort = this.sort;
    // Restore column visibility from localStorage
    const savedVisibility = localStorage.getItem('gfc_column_visibility');
    if (savedVisibility) {
      try {
        const parsed = JSON.parse(savedVisibility);
        this.columnVisibility = { ...this.columnVisibility, ...parsed };
        this.syncSelectedKeys();
        this.updateDisplayedColumns();
      } catch { /* ignore */ }
    }
    // Load persisted data from localStorage on startup
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          // Convert date strings back to Date objects
          parsed.forEach(record => {
            if (record.submittedAt) {
              const d = new Date(record.submittedAt);
              if (!isNaN(d.getTime())) record.submittedAt = d;
            }
          });
          this.dataSource.data = parsed;
          this.showToast(`${parsed.length} records loaded.`, 'info');
        }
      } catch { /* ignore corrupt data */ }
    }
  }

  // Column resizing: runs OUTSIDE Angular zone for performance
  initResize(event: MouseEvent, resizer: HTMLElement) {
    event.preventDefault();
    const th = resizer.closest('th');
    if (!th) return;
    const startX = event.pageX;
    const startWidth = th.offsetWidth;

    this.ngZone.runOutsideAngular(() => {
      const onMove = (e: MouseEvent) => {
        const newWidth = Math.max(60, startWidth + e.pageX - startX);
        th.style.width = newWidth + 'px';
        th.style.minWidth = newWidth + 'px';
      };
      const onUp = () => {
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
        document.body.style.cursor = '';
      };
      document.body.style.cursor = 'col-resize';
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    });
  }

  toggleReviewEdit(element: ApplicantRecord, field: 'isEditingReview' | 'isEditingHatyjaComments', event: MouseEvent) {
    event.stopPropagation();

    // First, clear any other editing states in all records to avoid multiple textareas
    this.dataSource.data.forEach(r => {
      r.isEditingReview = false;
      r.isEditingHatyjaComments = false;
    });

    this.editingReviewElement = element;

    // Set the specific editing flag
    if (field === 'isEditingReview') element.isEditingReview = true;
    else if (field === 'isEditingHatyjaComments') element.isEditingHatyjaComments = true;

    // Focus the textarea without jumping the scroll
    const container = (event.currentTarget as HTMLElement).closest('.review-container');
    setTimeout(() => {
      const textarea = container?.querySelector('textarea') as HTMLTextAreaElement;
      if (textarea) {
        textarea.focus({ preventScroll: true });
      }
    }, 50);
  }

  closeReviewEdit() {
    if (this.editingReviewElement) {
      this.editingReviewElement.isEditingReview = false;
      this.editingReviewElement.isEditingHatyjaComments = false;
    }
    this.editingReviewElement = null;
    this.saveToStorage();
  }

  applyFilter(event: Event) {
    const value = (event.target as HTMLInputElement).value.trim().toLowerCase();
    this.dataSource.filter = value;
    if (this.dataSource.paginator) {
      this.dataSource.paginator.firstPage();
    }
  }

  clearSearch(input: HTMLInputElement) {
    input.value = '';
    this.dataSource.filter = '';
    if (this.dataSource.paginator) {
      this.dataSource.paginator.firstPage();
    }
  }

  // ── Toast Notifications ──────────────────────────────────────
  toast: { message: string; type: 'success' | 'error' | 'info'; visible: boolean } = {
    message: '', type: 'info', visible: false
  };
  private toastTimer: any;

  showToast(message: string, type: 'success' | 'error' | 'info') {
    clearTimeout(this.toastTimer);
    this.toast = { message, type, visible: true };
    this.toastTimer = setTimeout(() => { this.toast.visible = false; }, 3000);
  }

  dismissToast() { this.toast.visible = false; }
  // ────────────────────────────────────────────────────────────────

  clearData() {
    localStorage.removeItem(STORAGE_KEY);
    this.dataSource.data = [];
    this.showToast('All data cleared from browser storage.', 'info');
  }

  saveToStorage() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(this.dataSource.data));
  }

  saveColumnVisibility() {
    localStorage.setItem('gfc_column_visibility', JSON.stringify(this.columnVisibility));
  }

  // Track the visibility of each column for selective export
  columnVisibility = {
    applicant: true,
    acronym: true,
    entityType: true,
    submittedAt: true,
    preScreening: true,
    profiles: true,
    country: true,
    address: true,
    nolStatus: true,
    hatyjaReviewComments: true,
    redFlags: true,
    passed: true,
    djResult: true,
    djReportNumber: true,
    djReportLink: true,
    djTruePositive: true,
    djFalsePositive: true,
    escalationRequired: true,
    hatyjaComments: true
  };

  updateDisplayedColumns() {
    this.displayedColumns = ['select', ...this.columnKeys.filter(key => (this.columnVisibility as any)[key])];
  }

  syncSelectedKeys() {
    this.selectedColumnKeys = this.columnKeys.filter(key => (this.columnVisibility as any)[key]);
  }

  onColumnSelectionChange() {
    // Reset all to false first
    Object.keys(this.columnVisibility).forEach(key => {
      (this.columnVisibility as any)[key] = false;
    });
    // Set selected to true
    this.selectedColumnKeys.forEach(key => {
      (this.columnVisibility as any)[key] = true;
    });
    this.updateDisplayedColumns();
    this.saveColumnVisibility();
  }

  removeColumn(key: string) {
    this.selectedColumnKeys = this.selectedColumnKeys.filter(k => k !== key);
    this.onColumnSelectionChange();
  }

  // Helper arrays for UI presentation
  columnKeys = [
    'applicant', 'acronym', 'entityType',
    'submittedAt', 'preScreening', 'profiles',
    'country', 'address', 'nolStatus', 'hatyjaReviewComments', 'redFlags', 'passed',
    'djResult', 'djReportNumber', 'djReportLink', 'djTruePositive', 'djFalsePositive', 'escalationRequired', 'hatyjaComments'
  ];
  columnNames: { [key: string]: string } = {
    applicant: 'Applicant',
    acronym: 'Acronym',
    entityType: 'Entity Type',
    submittedAt: 'Submitted At',
    preScreening: 'Pre-Screening',
    profiles: 'Profiles',
    country: 'Country',
    address: 'Address',
    nolStatus: 'NOL Status',
    hatyjaReviewComments: 'Hatyja Review comments',
    redFlags: 'Red Flags',
    passed: 'Passed',
    djResult: 'DJ-Result',
    djReportNumber: 'DJ report number',
    djReportLink: 'DJ report link',
    djTruePositive: 'DJ: True positive',
    djFalsePositive: 'DJ: False positive',
    escalationRequired: 'Escalation required to DRMC/Compliance?',
    hatyjaComments: 'Hatyja comments'
  };

  // Selection helpers
  isAllSelected() {
    const numSelected = this.selection.selected.length;
    const numRows = this.dataSource.data.length;
    return numSelected === numRows;
  }

  masterToggle() {
    this.isAllSelected() ?
      this.selection.clear() :
      this.dataSource.data.forEach(row => this.selection.select(row));
  }

  deleteSelectedRows() {
    const selected = this.selection.selected;
    if (selected.length === 0) return;

    const data = this.dataSource.data;
    const newData = data.filter(row => !this.selection.isSelected(row));
    this.dataSource.data = newData;
    this.selection.clear();
    this.saveToStorage();
    this.showToast(`${selected.length} records deleted.`, 'info');
  }

  // Remove individual deleteRow as it's replaced by the selection model

  exportToExcel(selectedOnly: boolean): void {
    let dataToExport: any[] = [];

    if (selectedOnly) {
      // Filter columns based on visibility
      dataToExport = this.dataSource.data.map(row => {
        const filteredRow: any = {};
        this.columnKeys.forEach(key => {
          if ((this.columnVisibility as any)[key]) {
            filteredRow[this.columnNames[key]] = (row as any)[key];
          }
        });
        return filteredRow;
      });
    } else {
      // Export all columns
      dataToExport = this.dataSource.data.map(row => {
        const fullRow: any = {};
        this.columnKeys.forEach(key => {
          fullRow[this.columnNames[key]] = (row as any)[key];
        });
        return fullRow;
      });
    }

    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook: XLSX.WorkBook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
    XLSX.writeFile(workbook, 'ApplicantData.xlsx');
    this.showToast('Excel exported successfully!', 'success');
  }

  copyTableToClipboard(selectedOnly: boolean): void {
    const data = this.dataSource.data;
    const columns = this.columnKeys.filter(key => (!selectedOnly || (this.columnVisibility as any)[key]));

    // Create headers row
    const headers = columns.map(key => this.columnNames[key]).join('\t');

    // Create data rowshatyja comments and reviews in blue
    const rows = data.map(row => {
      return columns.map(key => {
        let val = (row as any)[key];
        // Clean values for spreadsheet (remove newlines in reviews/addresses)
        if (typeof val === 'string') {
          val = val.replace(/\r?\n|\r/g, ' ');
        }
        return val || '';
      }).join('\t');
    });

    const tsv = [headers, ...rows].join('\n');

    navigator.clipboard.writeText(tsv).then(() => {
      this.showToast('Copied to clipboard! You can now paste into Excel.', 'success');
    }).catch(err => {
      this.showToast('Failed to copy to clipboard.', 'error');
      console.error('Clipboard error:', err);
    });
  }

  importExcel(event: any): void {
    const target: DataTransfer = <DataTransfer>(event.target);
    if (target.files.length !== 1) {
      this.showToast('Please select a single file.', 'error');
      return;
    }
    this.showToast('Importing data…', 'info');
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      const dataBuffer = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(dataBuffer, {
        type: 'array',
        cellDates: true,
        cellText: false,
        cellNF: true
      });

      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      // Use header: 1 to get an array of arrays instead of objects. This avoids header name mismatch issues.
      const importedData: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });

      const newRecords: ApplicantRecord[] = [];

      if (importedData.length > 0) {
        // En lugar de asumir que la fila 0 es la cabecera, buscamos cuál fila contiene las cabeceras reales
        let headerRowIndex = 0;
        let idxApp = -1, idxAcronym = -1, idxType = -1, idxCountry = -1, idxAddress = -1, idxNol = -1, idxReview = -1, idxFlags = -1, idxPassed = -1;
        let idxSubmitted = -1, idxPreScreen = -1, idxProfiles = -1;
        let idxDjResult = -1, idxDjNumber = -1, idxDjLink = -1, idxDjTrue = -1, idxDjFalse = -1, idxEscalation = -1, idxHatyjaExtra = -1;
        let foundHeaders = false;

        for (let i = 0; i < importedData.length && i < 10; i++) {
          const row = importedData[i];
          if (!row || !row.length) continue;

          let matches = 0;
          let tempIdxApp = -1, tempIdxAcronym = -1, tempIdxType = -1, tempIdxCountry = -1, tempIdxAddress = -1, tempIdxNol = -1, tempIdxReview = -1, tempIdxFlags = -1, tempIdxPassed = -1;
          let tempIdxSub = -1, tempIdxPre = -1, tempIdxProf = -1;
          let tempIdxDjRes = -1, tempIdxDjNum = -1, tempIdxDjLnk = -1, tempIdxDjT = -1, tempIdxDjF = -1, tempIdxEsc = -1, tempIdxHatX = -1;

          row.forEach((col: any, index: number) => {
            if (!col) return;
            const colName = String(col).toLowerCase().trim();
            if (colName.includes('applicant') || colName.includes('name')) { tempIdxApp = index; matches++; }
            else if (colName.includes('acronym') || colName.includes('short')) { tempIdxAcronym = index; matches++; }
            else if (colName.includes('entity') || colName.includes('type')) { tempIdxType = index; matches++; }
            else if (colName.includes('country') && !colName.includes('flag')) { tempIdxCountry = index; matches++; }
            else if (colName.includes('address') || colName.includes('location')) { tempIdxAddress = index; matches++; }
            else if (colName.includes('nol')) { tempIdxNol = index; matches++; }
            else if (colName === 'review' || (colName.includes('hatyja') && colName.includes('review'))) { tempIdxReview = index; matches++; }
            else if (colName.includes('red flag') || colName.includes('red-flag') || (colName.includes('flag') && !colName.includes('country'))) { tempIdxFlags = index; matches++; }
            else if (colName.includes('passed') || colName.includes('status')) { tempIdxPassed = index; matches++; }
            else if (colName.includes('submit') || colName.includes('subm') || colName.includes('date') || colName.includes('time') || colName.includes('create') || colName.includes('regist') || colName.includes('enviado') || colName.includes('fecha') || colName === 'at') { 
              tempIdxSub = index; 
              matches++; 
            }
            else if (colName.includes('screening') || colName.includes('pre-')) { tempIdxPre = index; matches++; }
            else if (colName.includes('profiles') || colName.includes('url') || colName.includes('link')) { tempIdxProf = index; matches++; }
            else if (colName.includes('dj-result') || colName.includes('dj result')) { tempIdxDjRes = index; matches++; }
            else if (colName.includes('report number') || colName.includes('dj report no')) { tempIdxDjNum = index; matches++; }
            else if (colName.includes('report link') || colName.includes('dj link')) { tempIdxDjLnk = index; matches++; }
            else if (colName.includes('true positive')) { tempIdxDjT = index; matches++; }
            else if (colName.includes('false positive')) { tempIdxDjF = index; matches++; }
            else if (colName.includes('escalation')) { tempIdxEsc = index; matches++; }
            else if (colName.includes('comments') && (colName.includes('hatyja') || colName.includes('extra') || colName.includes('review'))) { tempIdxHatX = index; matches++; }
          });

          if (matches >= 2) {
            console.log('--- HEADER FOUND AT ROW', i, '---');
            console.log('Indices:', { applicant: tempIdxApp, acronym: tempIdxAcronym, type: tempIdxType, country: tempIdxCountry, submitted: tempIdxSub });
            
            if (tempIdxSub === -1) console.warn('WARNING: Submitted column NOT found in this row headers.');
            
            idxApp = tempIdxApp !== -1 ? tempIdxApp : idxApp;
            idxAcronym = tempIdxAcronym !== -1 ? tempIdxAcronym : idxAcronym;
            idxType = tempIdxType !== -1 ? tempIdxType : idxType;
            idxCountry = tempIdxCountry !== -1 ? tempIdxCountry : idxCountry;
            idxAddress = tempIdxAddress;
            idxNol = tempIdxNol !== -1 ? tempIdxNol : idxNol;
            idxReview = tempIdxReview !== -1 ? tempIdxReview : idxReview;
            idxFlags = tempIdxFlags !== -1 ? tempIdxFlags : idxFlags;
            idxPassed = tempIdxPassed;
            idxSubmitted = tempIdxSub;
            idxPreScreen = tempIdxPre;
            idxProfiles = tempIdxProf;
            idxDjResult = tempIdxDjRes;
            idxDjNumber = tempIdxDjNum;
            idxDjLink = tempIdxDjLnk;
            idxDjTrue = tempIdxDjT;
            idxDjFalse = tempIdxDjF;
            idxEscalation = tempIdxEsc;
            idxHatyjaExtra = tempIdxHatX;
            headerRowIndex = i;
            foundHeaders = true;
            console.log('Detected headers at row', i, { idxApp, idxCountry, idxSubmitted, idxNol });
            break;
          }
        }

        // Start reading data from the row AFTER the header
        for (let i = headerRowIndex + 1; i < importedData.length; i++) {
          const row = importedData[i];
          if (row && row.length > 0) { // skip empty rows
            const hasData = row.some(cell => cell !== undefined && cell !== null && String(cell).trim() !== '');
            if (hasData) {
              const rawDate = idxSubmitted !== -1 ? row[idxSubmitted] : undefined;
              if (i === headerRowIndex + 1) {
                console.log('Sample data row:', row);
                console.log('Raw date value found:', rawDate, typeof rawDate);
              }

              newRecords.push({
                applicant: idxApp !== -1 && row[idxApp] !== undefined ? String(row[idxApp]) : '',
                acronym: idxAcronym !== -1 && row[idxAcronym] !== undefined ? String(row[idxAcronym]) : '',
                entityType: idxType !== -1 && row[idxType] !== undefined ? String(row[idxType]) : '',
                country: idxCountry !== -1 && row[idxCountry] !== undefined ? String(row[idxCountry]) : '',
                address: idxAddress !== -1 && row[idxAddress] !== undefined ? String(row[idxAddress]) : '',
                nolStatus: idxNol !== -1 && row[idxNol] !== undefined ? String(row[idxNol]) : '',
                hatyjaReviewComments: idxReview !== -1 && row[idxReview] !== undefined && String(row[idxReview]).trim() !== '' ? String(row[idxReview]) : '',
                redFlags: idxFlags !== -1 && row[idxFlags] !== undefined && String(row[idxFlags]).trim() !== '' ? String(row[idxFlags]) : '',
                passed: idxPassed !== -1 && row[idxPassed] !== undefined ? String(row[idxPassed]) : '',
                // Middle columns
                submittedAt: this.parseImportedDate(rawDate),
                preScreening: idxPreScreen !== -1 && row[idxPreScreen] !== undefined ? String(row[idxPreScreen]) : '',
                profiles: (() => {
                  const val = idxProfiles !== -1 && row[idxProfiles] !== undefined ? String(row[idxProfiles]).trim() : '';
                  return (val && !val.startsWith('http')) 
                    ? `https://partners.greenclimate.fund/pre-accreditation/${val}/staff/preview` 
                    : val;
                })(),
                // End columns
                djResult: idxDjResult !== -1 && row[idxDjResult] !== undefined ? String(row[idxDjResult]) : '',
                djReportNumber: idxDjNumber !== -1 && row[idxDjNumber] !== undefined ? String(row[idxDjNumber]) : '',
                djReportLink: idxDjLink !== -1 && row[idxDjLink] !== undefined ? String(row[idxDjLink]) : '',
                djTruePositive: idxDjTrue !== -1 && row[idxDjTrue] !== undefined ? String(row[idxDjTrue]) : '',
                djFalsePositive: idxDjFalse !== -1 && row[idxDjFalse] !== undefined ? String(row[idxDjFalse]) : '',
                escalationRequired: idxEscalation !== -1 && row[idxEscalation] !== undefined ? String(row[idxEscalation]) : '',
                hatyjaComments: idxHatyjaExtra !== -1 && row[idxHatyjaExtra] !== undefined ? String(row[idxHatyjaExtra]) : ''
              });
            }
          }
        }
      }

      this.dataSource.data = newRecords;
      localStorage.setItem(STORAGE_KEY, JSON.stringify(newRecords));
      console.log('Final Imported Records Sample:', newRecords[0]);
      if (this.paginator) {
        this.paginator.firstPage();
      }
      this.showToast(`${newRecords.length} records imported successfully!`, 'success');
      event.target.value = null;
    };
    reader.readAsArrayBuffer(target.files[0]);
  }

  private parseImportedDate(val: any): Date | undefined {
    if (!val) return undefined;
    if (val instanceof Date) return isNaN(val.getTime()) ? undefined : val;

    // Handle numbers (Excel serial date)
    if (typeof val === 'number') {
      // Excel uses 1900-01-01 as epoch. 25569 is the difference in days to Unix epoch.
      const d = new Date(Math.round((val - 25569) * 86400 * 1000));
      if (!isNaN(d.getTime())) return d;
    }

    // Handle strings
    if (typeof val === 'string' && val.trim() !== '') {
      // Try native JS parsing first (handles ISO and many common formats)
      let d = new Date(val);
      if (!isNaN(d.getTime())) {
        // Double check: if it parsed as a year like 0023, it might be a format error
        if (d.getFullYear() > 1900) return d;
      }

      // Try parsing common formats like DD/MM/YYYY or MM/DD/YYYY
      const parts = val.split(/[\/\-\.]/);
      if (parts.length === 3) {
        const p0 = Number(parts[0]);
        const p1 = Number(parts[1]);
        const p2 = Number(parts[2]);

        // YYYY/MM/DD
        if (parts[0].length === 4) {
          d = new Date(p0, p1 - 1, p2);
        }
        // Try MM/DD/YYYY (US, common in Excel) or DD/MM/YYYY
        else {
          // If middle value > 12, it must be the day: MM/DD/YYYY
          if (p1 > 12) {
            d = new Date(p2, p0 - 1, p1);
          } else {
            // Ambiguous case (4/10/2026), default to the one that gives a valid date
            // or just try common DD/MM/YYYY first, then MM/DD/YYYY
            d = new Date(p2, p1 - 1, p0); // DD/MM/YYYY
            if (isNaN(d.getTime())) d = new Date(p2, p0 - 1, p1); // MM/DD/YYYY
          }
        }
        if (!isNaN(d.getTime()) && d.getFullYear() > 1900) return d;
      }
    }

    return undefined;
  }
}

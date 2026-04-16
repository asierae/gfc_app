import { Component, ViewChild, AfterViewInit, NgZone } from '@angular/core';
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
import { MatTooltipModule } from '@angular/material/tooltip';
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
  entityTypeCounts: { name: string; count: number; pct: number }[];
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
    MatSortModule,
    MatFormFieldModule,
    MatInputModule,
    MatDatepickerModule,
    MatNativeDateModule,
    MatSelectModule,
    MatChipsModule,
    MatExpansionModule,
    MatTooltipModule
  ],
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements AfterViewInit {
  title = 'angular-excel-app';
  displayedColumns: string[] = [];
  dataSource = new MatTableDataSource<ApplicantRecord>(ELEMENT_DATA);
  selection = new SelectionModel<ApplicantRecord>(true, []);
  
  // Filtering properties
  filterText: string = '';
  filterStartDate?: Date | null;
  filterEndDate?: Date | null;
  filterRegion: string = 'All';
  filterEntityType: string = 'All';
  filterStatus: string = 'All';
  readonly statusOptions: string[] = ['All', 'Accepted', 'Rejected', 'Pending'];
  statusOptionsWithCount: { value: string; label: string }[] = [];
  readonly entityTypeOptions: string[] = [
    'All',
    'UN System Entity',
    'Public/Government-Controlled',
    'Public/Government',
    'Private Sector',
    'Multilateral Organization',
    'Unknown'
  ];
  entityTypeOptionsWithCount: { value: string; label: string }[] = [];
  readonly regionOptions: string[] = [
    'All',
    'DAFR',
    'DAPAC',
    'DECM',
    'DLAC',
    'INTERNATIONAL'
  ];
  regionOptionsWithCount: { value: string; label: string }[] = [];
  private readonly regionCountryMap: Record<string, Set<string>> = {
    DAPAC: new Set([
      'afghanistan', 'bangladesh', 'bhutan', 'china', 'hong kong', 'india', 'japan', 'kazakhstan',
      'kyrgyzstan', 'laos', 'macau', 'maldives', 'mongolia', 'nepal', 'north korea', 'pakistan',
      'south korea', 'sri lanka', 'taiwan', 'tajikistan', 'turkmenistan', 'uzbekistan', 'vietnam',
      'viet nam',
      "lao people's democratic", "lao people's democratic republic", "lao people's democratic republic (the)", 'lao pdr',
      'fiji', 'kiribati', 'marshall islands', 'micronesia', 'nauru', 'palau', 'papua new guinea',
      'samoa', 'solomon islands', 'timor leste', 'timor-leste', 'tonga', 'tuvalu', 'vanuatu'
      , 'brunei', 'cambodia', 'indonesia', 'malaysia', 'myanmar', 'philippines', 'singapore',
      'thailand'
    ]),
    DAFR: new Set([
      'algeria', 'benin', 'burkina faso', 'cabo verde', 'cape verde', 'cameroon',
      'central african republic', 'chad', 'congo', 'democratic republic of the congo',
      'egypt', 'equatorial guinea', 'gabon', 'gambia', 'ghana', 'guinea', 'guinea-bissau',
      'ivory coast', 'cote d ivoire', "cote d'ivoire", 'liberia', 'libya', 'mali', 'mauritania',
      'morocco', 'niger', 'nigeria', 'senegal', 'sierra leone', 'sudan', 'togo', 'tunisia',
      'angola', 'botswana', 'burundi', 'comoros', 'djibouti', 'eritrea', 'eswatini', 'swaziland',
      'ethiopia', 'kenya', 'lesotho', 'madagascar', 'malawi', 'mauritius', 'mozambique',
      'namibia', 'rwanda', 'seychelles', 'somalia', 'south africa', 'south sudan', 'tanzania',
      'uganda', 'zambia', 'zimbabwe'
    ]),
    DECM: new Set([
      'albania', 'andorra', 'armenia', 'austria', 'azerbaijan', 'belarus', 'belgium',
      'bosnia and herzegovina', 'bulgaria', 'croatia', 'cyprus', 'czech republic', 'czechia',
      'denmark', 'estonia', 'finland', 'france', 'georgia', 'germany', 'greece', 'hungary',
      'iceland', 'ireland', 'italy', 'kosovo', 'latvia', 'liechtenstein', 'lithuania',
      'luxembourg', 'malta', 'moldova', 'monaco', 'montenegro', 'netherlands', 'netherlands (the)',
      'north macedonia', 'norway', 'poland', 'portugal', 'romania', 'russia', 'san marino',
      'serbia', 'slovakia', 'slovenia', 'spain', 'sweden', 'switzerland', 'uk', 'united kingdom',
      'ukraine', 'vatican city',
      'bahrain', 'iraq', 'israel', 'jordan', 'kuwait', 'lebanon', 'oman', 'palestine', 'qatar',
      'saudi arabia', 'syria', 'turkey', 'united arab emirates', 'uae', 'yemen', 'iran'
    ]),
    DLAC: new Set([
      'antigua and barbuda', 'bahamas', 'barbados', 'belize', 'dominica', 'grenada', 'guyana',
      'haiti', 'jamaica', 'saint kitts and nevis', 'saint lucia', 'saint vincent and the grenadines',
      'suriname', 'trinidad and tobago',
      'argentina', 'bolivia', 'brazil', 'chile', 'colombia', 'costa rica', 'cuba',
      'dominican republic', 'ecuador', 'el salvador', 'guatemala', 'honduras', 'mexico',
      'nicaragua', 'panama', 'paraguay', 'peru', 'uruguay', 'venezuela'
    ])
  };
  private readonly countryToRegionMap = new Map<string, string>();

  @ViewChild(MatPaginator) paginator!: MatPaginator;
  @ViewChild(MatSort) sort!: MatSort;
  editingReviewElement: ApplicantRecord | null = null;
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
        topCountriesRedFlags: [],
        entityTypeCounts: []
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
      acceptedPct: Number(((accepted / total) * 100).toFixed(3)),
      rejectedPct: Number(((rejected / total) * 100).toFixed(3)),
      pendingPct: Number(((pending / total) * 100).toFixed(3)),
      redFlagsPct: Number(((redFlags / total) * 100).toFixed(3)),
      topCountriesRedFlags: this.getTopCountriesRedFlags(data),
      entityTypeCounts: this.getEntityTypeCounts(data, total)
    };
  }

  readonly entityTypeColorMap: Record<string, string> = {
    'UN System Entity': '#3b82f6',
    'Public/Government-Controlled': '#8b5cf6',
    'Public/Government': '#06b6d4',
    'Private Sector': '#f59e0b',
    'Multilateral Organization': '#10b981',
    'Unknown': '#94a3b8'
  };

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
    this.buildCountryToRegionMap();
    this.syncSelectedKeys();
    this.updateDisplayedColumns();
    this.deferFilterCountRefresh();
  }

  ngAfterViewInit() {
    this.dataSource.paginator = this.paginator;
    this.dataSource.sort = this.sort;
    
    // Set up custom sorting logic for dates vs text
    this.dataSource.sortingDataAccessor = (item: ApplicantRecord, property: string) => {
      switch (property) {
        case 'submittedAt':
          return item.submittedAt ? new Date(item.submittedAt).getTime() : 0;
        default:
          const val = (item as any)[property];
          return typeof val === 'string' ? val.toLowerCase() : val || '';
      }
    };
    
    // Set up custom filter predicate
    this.dataSource.filterPredicate = (data: ApplicantRecord, filter: string) => {
      const defaultTerms = { text: '', region: 'All', entityType: 'All', status: 'All', start: null, end: null };
      const searchTerms: { text: string; region: string; entityType: string; status: string; start?: Date | null; end?: Date | null } =
        (typeof filter === 'string' && filter.trim().startsWith('{'))
          ? JSON.parse(filter)
          : defaultTerms;
      
      // 1. Text Search (over specific columns)
      const dataStr = (
        (data.applicant || '') + 
        (data.acronym || '') + 
        (data.country || '') + 
        (data.nolStatus || '')
      ).toLowerCase();
      
      const matchesSearch = dataStr.includes((searchTerms.text || '').toLowerCase());

      // 1.5. Region Filter (computed from country field)
      const rowRegion = this.getRegionByCountry(data.country);
      const matchesRegion = this.matchesRegionSelection(rowRegion, searchTerms.region);
      
      // 2. Date Range Filter
      let matchesDate = true;
      if (searchTerms.start || searchTerms.end) {
        if (!data.submittedAt) {
          matchesDate = false;
        } else {
          const subDate = new Date(data.submittedAt);
          subDate.setHours(0,0,0,0);
          
          if (searchTerms.start) {
            const start = new Date(searchTerms.start);
            start.setHours(0,0,0,0);
            if (subDate < start) matchesDate = false;
          }
          if (searchTerms.end) {
            const end = new Date(searchTerms.end);
            end.setHours(0,0,0,0);
            if (subDate > end) matchesDate = false;
          }
        }
      }
      
      // 3. Entity Type Filter
      const matchesEntityType = this.matchesEntityTypeSelection(data.entityType, searchTerms.entityType);

      // 4. Status Filter (Passed column)
      let matchesStatus = true;
      if (searchTerms.status && searchTerms.status !== 'All') {
        if (searchTerms.status === 'Pending') {
          matchesStatus = !data.passed || (data.passed !== 'Accepted' && data.passed !== 'Rejected');
        } else {
          matchesStatus = data.passed === searchTerms.status;
        }
      }

      return matchesSearch && matchesRegion && matchesDate && matchesEntityType && matchesStatus;
    };

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
          this.deferFilterCountRefresh();
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
    this.filterText = (event.target as HTMLInputElement).value.trim().toLowerCase();
    this.updateFilter();
  }

  updateFilter() {
    const filterValue = {
      text: this.filterText,
      region: this.filterRegion,
      entityType: this.filterEntityType,
      status: this.filterStatus,
      start: this.filterStartDate,
      end: this.filterEndDate
    };
    this.dataSource.filter = JSON.stringify(filterValue);
    
    if (this.dataSource.paginator) {
      this.dataSource.paginator.firstPage();
    }
  }

  clearSearch(input: HTMLInputElement) {
    input.value = '';
    this.filterText = '';
    this.updateFilter();
  }

  clearDateFilter() {
    this.filterStartDate = null;
    this.filterEndDate = null;
    this.updateFilter();
  }

  onRegionFilterChange() {
    this.updateFilter();
  }

  onEntityTypeFilterChange() {
    this.updateFilter();
  }

  onStatusFilterChange() {
    this.updateFilter();
  }

  private getRegionByCountry(country?: string): string {
    const normalized = this.normalizeCountryName(country);
    if (!normalized) return 'INTERNATIONAL';
    return this.countryToRegionMap.get(normalized) || 'INTERNATIONAL';
  }

  private matchesRegionSelection(rowRegion: string, selectedRegion?: string): boolean {
    if (!selectedRegion || selectedRegion === 'All') return true;
    if (rowRegion === selectedRegion) return true;
    return false;
  }

  private normalizeCountryName(country?: string): string {
    if (!country) return '';
    return country
      .toLowerCase()
      .trim()
      .replace(/\(the\)/g, '')
      .replace(/[.,]/g, '')
      .replace(/\s+/g, ' ');
  }

  private getRegionCounts(): Record<string, number> {
    const rowRegions = this.dataSource.data.map(record => this.getRegionByCountry(record.country));
    const counts: Record<string, number> = {};

    this.regionOptions.forEach(regionOption => {
      counts[regionOption] = rowRegions.filter(rowRegion =>
        this.matchesRegionSelection(rowRegion, regionOption)
      ).length;
    });

    return counts;
  }

  private buildCountryToRegionMap() {
    this.countryToRegionMap.clear();
    Object.entries(this.regionCountryMap).forEach(([region, countries]) => {
      countries.forEach(country => {
        this.countryToRegionMap.set(this.normalizeCountryName(country), region);
      });
    });
  }

  private refreshRegionOptionsWithCount() {
    const counts = this.getRegionCounts();
    this.regionOptionsWithCount = this.regionOptions.map(region => ({
      value: region,
      label: `${region} (${counts[region] ?? 0})`
    }));
  }

  private normalizeEntityType(entityType?: string): string {
    if (!entityType) return 'Unknown';
    const val = entityType.toLowerCase().trim();
    if (val.includes('un') && val.includes('system')) return 'UN System Entity';
    if (val.includes('un-system')) return 'UN System Entity';
    if (val.includes('government-controlled') || val.includes('government controlled')) return 'Public/Government-Controlled';
    if (val.includes('public') && val.includes('government') && val.includes('controlled')) return 'Public/Government-Controlled';
    if (val.includes('public') && val.includes('government')) return 'Public/Government';
    if (val === 'public/government' || val === 'public-government') return 'Public/Government';
    if (val.includes('private') || val.includes('sector')) return 'Private Sector';
    if (val.includes('multilateral') || val.includes('organization')) return 'Multilateral Organization';
    if (val.includes('non-profit') || val.includes('nonprofit')) return 'Multilateral Organization';
    if (val.includes('corporation') || val.includes('llc') || val.includes('partnership')) return 'Private Sector';
    return 'Unknown';
  }

  private matchesEntityTypeSelection(entityType?: string, selectedType?: string): boolean {
    if (!selectedType || selectedType === 'All') return true;
    return this.normalizeEntityType(entityType) === selectedType;
  }

  private getEntityTypeCounts(data: ApplicantRecord[], total: number): { name: string; count: number; pct: number }[] {
    const counts: Record<string, number> = {};
    data.forEach(d => {
      const normalized = this.normalizeEntityType(d.entityType);
      counts[normalized] = (counts[normalized] || 0) + 1;
    });
    return Object.keys(counts)
      .map(name => ({ name, count: counts[name], pct: total > 0 ? Number(((counts[name] / total) * 100).toFixed(1)) : 0 }))
      .sort((a, b) => b.count - a.count);
  }

  private getEntityTypeCounts2(): Record<string, number> {
    const counts: Record<string, number> = {};
    this.entityTypeOptions.forEach(opt => { counts[opt] = 0; });
    this.dataSource.data.forEach(record => {
      const normalized = this.normalizeEntityType(record.entityType);
      counts[normalized] = (counts[normalized] || 0) + 1;
      counts['All'] = (counts['All'] || 0) + 1;
    });
    return counts;
  }

  private refreshEntityTypeOptionsWithCount() {
    const counts = this.getEntityTypeCounts2();
    this.entityTypeOptionsWithCount = this.entityTypeOptions.map(opt => ({
      value: opt,
      label: `${opt} (${counts[opt] ?? 0})`
    }));
  }

  private refreshStatusOptionsWithCount() {
    const data = this.dataSource.data;
    const counts: Record<string, number> = { All: data.length };
    data.forEach(d => {
      if (d.passed === 'Accepted') counts['Accepted'] = (counts['Accepted'] || 0) + 1;
      else if (d.passed === 'Rejected') counts['Rejected'] = (counts['Rejected'] || 0) + 1;
      else counts['Pending'] = (counts['Pending'] || 0) + 1;
    });
    this.statusOptionsWithCount = this.statusOptions.map(opt => ({
      value: opt,
      label: `${opt} (${counts[opt] ?? 0})`
    }));
  }

  private deferFilterCountRefresh() {
    setTimeout(() => {
      this.refreshRegionOptionsWithCount();
      this.refreshEntityTypeOptionsWithCount();
      this.refreshStatusOptionsWithCount();
    }, 0);
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
    this.deferFilterCountRefresh();
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
    this.deferFilterCountRefresh();
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
      this.deferFilterCountRefresh();
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
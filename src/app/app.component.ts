import { Component, ViewChild, AfterViewInit, OnDestroy, NgZone } from '@angular/core';
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
import { MatMenuModule } from '@angular/material/menu';
import * as XLSX from 'xlsx';
import { FirestoreService } from './firestore.service';
import { SkillsService } from './services/skills.service';
import { RiskService } from './services/risk.service';
import { InvestigationSkill } from './config/skills.config';
import {
  REGION_OPTIONS,
  buildCountryToRegionMap,
  normalizeCountryName,
  getRegionByCountry,
} from './config/countries.config';

export interface ApplicantRecord {
  id?: string;
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
  window?: number | null;
  invited?: Date | string | null;
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
  emailCommunications?: string;
  // DRMC/Compliance QA columns
  drmcCompliance1?: string;
  drmcCompliance2?: string;
  // Risk analysis
  riskPercent?: number | null;
  riskReasons?: string;
  isCalculatingRisk?: boolean;
  // UI flags
  isEditingReview?: boolean;
  isEditingHatyjaComments?: boolean;
  isEditingEmailCommunications?: boolean;
  isEditingRiskReasons?: boolean;
  archived?: boolean;
}

export interface DashboardStats {
  total: number;
  countries: number;
  passed: number;
  failed: number;
  invited: number;
  pending: number;
  redFlags: number;
  passedPct: number;
  failedPct: number;
  invitedPct: number;
  pendingPct: number;
  redFlagsPct: number;
  topCountriesRedFlags: { name: string; count: number; avgRisk: number }[];
  entityTypeCounts: { name: string; count: number; pct: number }[];
}

const ELEMENT_DATA: ApplicantRecord[] = [];

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
    MatTooltipModule,
    MatMenuModule,
  ],
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent implements AfterViewInit, OnDestroy {
  title = 'angular-excel-app';
  displayedColumns: string[] = [];
  dataSource = new MatTableDataSource<ApplicantRecord>(ELEMENT_DATA);
  selection = new SelectionModel<ApplicantRecord>(true, []);

  // Filtering properties
  filterText: string = '';
  filterStartDate?: Date | null;
  filterEndDate?: Date | null;
  filterWindow: string = 'All';
  filterRegion: string = 'All';
  filterEntityType: string = 'All';
  filterStatus: string = 'All';
  showArchived = false;
  readonly statusOptions: string[] = ['All', 'Pending', 'Passed', 'Failed', 'Invited'];
  readonly windowOptions: string[] = ['All', '1', '2', '3'];
  windowOptionsWithCount: { value: string; label: string }[] = [];
  private readonly validPassedValues = new Set(['Pending', 'Passed', 'Failed', 'Invited']);

  // --- Risk Investigation Skills ---
  investigationSkills: InvestigationSkill[] = [];

  // Gemini API configuration
  private geminiApiKey: string = '';
  private readonly GEMINI_KEY_STORAGE = 'gfc_gemini_api_key';

  // ID generation helper
  private generateApplicantId(): string {
    return 'app_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
  }
  statusOptionsWithCount: { value: string; label: string }[] = [];
  readonly entityTypeOptions: string[] = [
    'All',
    'UN System Entity',
    'Public/Government-Controlled',
    'Public/Government',
    'Private Sector',
    'Multilateral Organization',
    'Unknown',
  ];
  entityTypeOptionsWithCount: { value: string; label: string }[] = [];
  readonly regionOptions: string[] = [...REGION_OPTIONS];
  regionOptionsWithCount: { value: string; label: string }[] = [];
  private readonly countryToRegionMap = buildCountryToRegionMap();

  @ViewChild(MatPaginator) paginator!: MatPaginator;
  @ViewChild(MatSort) sort!: MatSort;
  editingReviewElement: ApplicantRecord | null = null;
  selectedColumnKeys: string[] = [];
  firestoreConnected = false;
  private lastFirestoreErrorToastAt = 0;
  private readonly onBrowserOnline = () => {
    this.setupRealtimeSubscription();
    void this.flushPendingLocalChanges(true);
  };
  private readonly onBrowserOffline = () =>
    this.setFirestoreConnected(
      false,
      'Connection to Firebase lost. Changes are saved locally only.',
    );

  private get activeRecords(): ApplicantRecord[] {
    return this.dataSource.data.filter((r) => !r.archived);
  }

  get stats(): DashboardStats {
    const data = this.activeRecords;
    const total = data.length;
    if (total === 0) {
      return {
        total: 0,
        countries: 0,
        passed: 0,
        failed: 0,
        invited: 0,
        pending: 0,
        redFlags: 0,
        passedPct: 0,
        failedPct: 0,
        invitedPct: 0,
        pendingPct: 0,
        redFlagsPct: 0,
        topCountriesRedFlags: [],
        entityTypeCounts: [],
      };
    }

    const countries = new Set(data.map((d) => d.country).filter((c) => !!c)).size;
    const passed = data.filter((d) => d.passed === 'Passed').length;
    const failed = data.filter((d) => d.passed === 'Failed').length;
    const invited = data.filter((d) => d.passed === 'Invited').length;
    const pending = data.filter((d) => !d.passed || d.passed === 'Pending').length;
    const redFlags = data.filter((d) => d.redFlags && d.redFlags !== 'None').length;

    return {
      total,
      countries,
      passed,
      failed,
      invited,
      pending,
      redFlags,
      passedPct: Number(((passed / total) * 100).toFixed(3)),
      failedPct: Number(((failed / total) * 100).toFixed(3)),
      invitedPct: Number(((invited / total) * 100).toFixed(3)),
      pendingPct: Number(((pending / total) * 100).toFixed(3)),
      redFlagsPct: Number(((redFlags / total) * 100).toFixed(3)),
      topCountriesRedFlags: this.getTopCountriesRedFlags(data),
      entityTypeCounts: this.getEntityTypeCounts(data, total),
    };
  }

  readonly entityTypeColorMap: Record<string, string> = {
    'UN System Entity': '#3b82f6',
    'Public/Government-Controlled': '#8b5cf6',
    'Public/Government': '#06b6d4',
    'Private Sector': '#f59e0b',
    'Multilateral Organization': '#10b981',
    Unknown: '#94a3b8',
  };

  private getTopCountriesRedFlags(data: ApplicantRecord[]) {
    const grouped: Record<string, { total: number; sum: number }> = {};
    data
      .filter((d) => d.country && d.riskPercent != null)
      .forEach((d) => {
        if (!grouped[d.country]) grouped[d.country] = { total: 0, sum: 0 };
        grouped[d.country].total++;
        grouped[d.country].sum += d.riskPercent!;
      });

    return Object.keys(grouped)
      .map((name) => ({
        name,
        count: grouped[name].total,
        avgRisk: Math.round(grouped[name].sum / grouped[name].total),
      }))
      .sort((a, b) => b.avgRisk - a.avgRisk)
      .slice(0, 5);
  }

  constructor(
    private ngZone: NgZone,
    private firestoreService: FirestoreService,
    private skillsService: SkillsService,
    private riskService: RiskService,
  ) {
    this.syncSelectedKeys();
    this.updateDisplayedColumns();
    this.deferFilterCountRefresh();
    this.geminiApiKey = localStorage.getItem(this.GEMINI_KEY_STORAGE) || '';
    this.loadSkills();
  }

  ngAfterViewInit() {
    this.dataSource.paginator = this.paginator;
    this.dataSource.sort = this.sort;

    // Default sort: Submitted At ascending (earliest first)
    this.sort.active = 'submittedAt';
    this.sort.direction = 'asc';
    this.sort.sortChange.emit({ active: 'submittedAt', direction: 'asc' });

    // Set up custom sorting logic for dates vs text
    this.dataSource.sortingDataAccessor = (item: ApplicantRecord, property: string) => {
      switch (property) {
        case 'submittedAt':
          return item.submittedAt ? new Date(item.submittedAt).getTime() : 0;
        case 'invited':
          return item.invited ? new Date(item.invited).getTime() : 0;
        default:
          const val = (item as any)[property];
          return typeof val === 'string' ? val.toLowerCase() : val || '';
      }
    };

    // Set up custom filter predicate
    this.dataSource.filterPredicate = (data: ApplicantRecord, filter: string) => {
      const defaultTerms = {
        text: '',
        region: 'All',
        entityType: 'All',
        status: 'All',
        window: 'All',
        start: null,
        end: null,
        showArchived: false,
      };
      const searchTerms: {
        text: string;
        region: string;
        entityType: string;
        status: string;
        window?: string;
        start?: Date | null;
        end?: Date | null;
        showArchived?: boolean;
      } =
        typeof filter === 'string' && filter.trim().startsWith('{')
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
          subDate.setHours(0, 0, 0, 0);

          if (searchTerms.start) {
            const start = new Date(searchTerms.start);
            start.setHours(0, 0, 0, 0);
            if (subDate < start) matchesDate = false;
          }
          if (searchTerms.end) {
            const end = new Date(searchTerms.end);
            end.setHours(0, 0, 0, 0);
            if (subDate > end) matchesDate = false;
          }
        }
      }

      // 3. Entity Type Filter
      const matchesEntityType = this.matchesEntityTypeSelection(
        data.entityType,
        searchTerms.entityType,
      );

      // 4. Status Filter (Passed column)
      let matchesStatus = true;
      if (searchTerms.status && searchTerms.status !== 'All') {
        if (searchTerms.status === 'Pending') {
          matchesStatus = !data.passed || data.passed === 'Pending';
        } else {
          matchesStatus = data.passed === searchTerms.status;
        }
      }

      // 5. Window dropdown filter (exact match)
      let matchesWindow = true;
      if (searchTerms.window && searchTerms.window !== 'All') {
        const w =
          data.window !== undefined && data.window !== null ? Number((data as any).window) : NaN;
        matchesWindow = !isNaN(w) && w === parseInt(searchTerms.window, 10);
      }

      const isArchived = !!data.archived;
      const showArchived = searchTerms.showArchived === true;
      const matchesArchived = showArchived ? isArchived : !isArchived;

      return (
        matchesSearch &&
        matchesRegion &&
        matchesDate &&
        matchesEntityType &&
        matchesStatus &&
        matchesWindow &&
        matchesArchived
      );
    };

    // Restore column visibility from localStorage, then Firestore
    const savedVisibility = localStorage.getItem('gfc_column_visibility');
    if (savedVisibility) {
      try {
        const parsed = JSON.parse(savedVisibility);
        this.columnVisibility = { ...this.columnVisibility, ...parsed };
        this.syncSelectedKeys();
        this.updateDisplayedColumns();
      } catch {
        /* ignore */
      }
    }
    this.firestoreService.loadColumnVisibility().then((parsed) => {
      if (parsed) {
        this.ngZone.run(() => {
          this.columnVisibility = { ...this.columnVisibility, ...parsed };
          this.syncSelectedKeys();
          this.updateDisplayedColumns();
          localStorage.setItem('gfc_column_visibility', JSON.stringify(this.columnVisibility));
        });
      }
    });
    // Load from localStorage first (instant), then sync from Firestore
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          this.hydrateRecords(parsed);
          this.dataSource.data = parsed;
          this.restoreDirtyApplicants();
          this.deferFilterCountRefresh();
          this.updateFilter();
          this.showToast(`${parsed.length} records loaded (local cache).`, 'info');
        }
      } catch {
        /* ignore corrupt data */
      }
    } else {
      this.updateFilter();
    }

    // Then load from Firestore (source of truth)
    this.loadFromFirestore();

    window.addEventListener('online', this.onBrowserOnline);
    window.addEventListener('offline', this.onBrowserOffline);
  }

  ngOnDestroy() {
    window.removeEventListener('online', this.onBrowserOnline);
    window.removeEventListener('offline', this.onBrowserOffline);
    if (this.unsubscribeFromApplicants) {
      this.unsubscribeFromApplicants();
    }
  }

  private setFirestoreConnected(connected: boolean, errorMessage?: string) {
    this.ngZone.run(() => {
      this.firestoreConnected = connected;
      if (!connected && errorMessage) {
        const now = Date.now();
        if (now - this.lastFirestoreErrorToastAt > 3000) {
          this.lastFirestoreErrorToastAt = now;
          this.showToast(errorMessage, 'error');
        }
      }
    });
  }

  private hydrateRecords(records: any[]) {
    records.forEach((record) => {
      if (record.submittedAt) {
        const d = new Date(record.submittedAt);
        if (!isNaN(d.getTime())) record.submittedAt = d;
      }
      if (record.invited) {
        const di = new Date(record.invited);
        if (!isNaN(di.getTime())) record.invited = di;
      }
    });
  }

  /** Merge Firestore updates in place so pagination and sort position are preserved. */
  private mergeApplicantsFromFirestore(applicants: ApplicantRecord[]) {
    if (this.importInProgress) return;

    const incomingById = new Map(applicants.filter((a) => a.id).map((a) => [a.id!, a]));
    const firestoreIds = new Set(incomingById.keys());
    let structureChanged = false;

    const kept: ApplicantRecord[] = [];
    for (const record of this.dataSource.data) {
      if (record.id && incomingById.has(record.id)) {
        const incoming = incomingById.get(record.id)!;
        if (this.shouldKeepLocalVersion(record.id, (incoming as any).updatedAt)) {
          kept.push(record);
        } else {
          Object.assign(record, incoming);
          this.dirtyApplicantIds.delete(record.id);
          kept.push(record);
        }
        incomingById.delete(record.id);
      } else if (record.id && !firestoreIds.has(record.id)) {
        kept.push(record);
        structureChanged = true;
      } else {
        kept.push(record);
      }
    }

    if (incomingById.size > 0) {
      structureChanged = true;
      kept.push(...incomingById.values());
    }

    if (structureChanged) {
      this.dataSource.data = kept;
    } else if (kept.length > 0) {
      this.dataSource.data = [...kept];
    }

    this.persistDirtyApplicants();
  }

  private shouldKeepLocalVersion(recordId: string, remoteUpdatedAt?: string): boolean {
    const localDirtyAt = this.dirtyApplicantIds.get(recordId);
    if (!localDirtyAt) return false;
    if (!remoteUpdatedAt) return true;
    return !this.isRemoteNewerThanLocal(remoteUpdatedAt, localDirtyAt);
  }

  private restoreDirtyApplicants() {
    const raw = localStorage.getItem(this.DIRTY_APPLICANTS_KEY);
    if (!raw) return;
    try {
      const entries: [string, string][] = JSON.parse(raw);
      this.dirtyApplicantIds = new Map(entries);
    } catch {
      /* ignore corrupt data */
    }
  }

  private persistDirtyApplicants() {
    localStorage.setItem(
      this.DIRTY_APPLICANTS_KEY,
      JSON.stringify([...this.dirtyApplicantIds.entries()]),
    );
  }

  private readonly LOCAL_TIMESTAMP_KEY = 'gfc_data_updated_at';
  private readonly DIRTY_APPLICANTS_KEY = 'gfc_dirty_applicants';
  private dirtyApplicantIds = new Map<string, string>();
  private importInProgress = false;
  private unsubscribeFromApplicants: (() => void) | null = null;
  private lastAppliedFilter = '';

  private async loadFromFirestore() {
    try {
      const migrated = await this.firestoreService.migrateFromLegacy(() =>
        this.generateApplicantId(),
      );
      if (migrated > 0) {
        this.showToast(`Migrated ${migrated} records to new format.`, 'success');
      }

      const applicants = await this.firestoreService.loadAllApplicants();
      this.setFirestoreConnected(true);

      if (applicants.length > 0) {
        this.hydrateRecords(applicants);
        this.ngZone.run(() => {
          if (this.dataSource.data.length > 0) {
            this.mergeApplicantsFromFirestore(applicants);
          } else {
            this.dataSource.data = applicants;
          }
          this.deferFilterCountRefresh();
          this.updateFilter();
          localStorage.setItem(STORAGE_KEY, JSON.stringify(this.dataSource.data));
          this.showToast(`${applicants.length} records synced from cloud.`, 'success');
        });
      }
    } catch (err) {
      console.error('Firestore: initial load failed', err);
      this.setFirestoreConnected(false, 'Could not connect to Firebase.');
    }

    this.setupRealtimeSubscription();
    void this.flushPendingLocalChanges();
  }

  private setupRealtimeSubscription() {
    // Unsubscribe from previous subscription if exists
    if (this.unsubscribeFromApplicants) {
      this.unsubscribeFromApplicants();
    }

    this.unsubscribeFromApplicants = this.firestoreService.subscribeToApplicants(
      (applicants, fromServer) => {
        applicants.forEach((a) => {
          if (!a.id) a.id = this.generateApplicantId();
        });

        this.hydrateRecords(applicants);
        this.ngZone.run(() => {
          this.mergeApplicantsFromFirestore(applicants);
          this.deferFilterCountRefresh();
          localStorage.setItem(STORAGE_KEY, JSON.stringify(this.dataSource.data));
        });

        if (fromServer || navigator.onLine) {
          this.setFirestoreConnected(true);
        }
      },
      () => this.setFirestoreConnected(false, 'Firebase sync interrupted.'),
    );
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

  private preEditSortState: { active: string; direction: string } | null = null;
  private preEditDataSourceSort: MatSort | null | undefined = null;

  toggleReviewEdit(
    element: ApplicantRecord,
    field:
      | 'isEditingReview'
      | 'isEditingHatyjaComments'
      | 'isEditingEmailCommunications'
      | 'isEditingRiskReasons',
    event: MouseEvent,
  ) {
    event.stopPropagation();

    // First, clear any other editing states in all records to avoid multiple textareas
    this.dataSource.data.forEach((r) => {
      r.isEditingReview = false;
      r.isEditingHatyjaComments = false;
      r.isEditingEmailCommunications = false;
      r.isEditingRiskReasons = false;
    });

    this.editingReviewElement = element;

    // Set the specific editing flag
    if (field === 'isEditingReview') element.isEditingReview = true;
    else if (field === 'isEditingHatyjaComments') element.isEditingHatyjaComments = true;
    else if (field === 'isEditingEmailCommunications') element.isEditingEmailCommunications = true;
    else if (field === 'isEditingRiskReasons') element.isEditingRiskReasons = true;

    // Capture rect NOW — event.currentTarget is lost inside setTimeout
    const triggerEl = event.currentTarget as HTMLElement;
    const triggerRect = triggerEl.getBoundingClientRect();
    const container = triggerEl.closest('.review-container');
    setTimeout(() => {
      const textarea = container?.querySelector('textarea') as HTMLTextAreaElement;
      if (textarea) {
        const margin = 12;
        const minHeight = 150;
        const maxPreferredHeight = 320;
        const gap = 8;
        const desiredWidth = Math.min(420, window.innerWidth - margin * 2);

        const availableAbove = Math.max(0, triggerRect.top - margin - gap);
        const availableBelow = Math.max(0, window.innerHeight - triggerRect.bottom - margin - gap);
        const isTopZone = triggerRect.top < window.innerHeight * 0.45;
        let placeBelow = isTopZone;
        if (!isTopZone) {
          if (availableBelow >= minHeight) {
            placeBelow = true;
          } else if (availableAbove >= minHeight) {
            placeBelow = false;
          } else {
            placeBelow = availableBelow >= availableAbove;
          }
        }

        const availableOnSide = placeBelow ? availableBelow : availableAbove;
        const popupMaxHeight = Math.max(
          minHeight,
          Math.min(maxPreferredHeight, availableOnSide || minHeight),
        );

        textarea.style.maxHeight = `${popupMaxHeight}px`;
        textarea.style.overflowY = 'auto';

        let left = triggerRect.right - desiredWidth;
        left = Math.max(margin, Math.min(left, window.innerWidth - desiredWidth - margin));

        textarea.style.position = 'fixed';
        textarea.style.bottom = 'auto';
        textarea.style.right = 'auto';
        textarea.style.left = `${left}px`;
        textarea.style.width = `${desiredWidth}px`;
        textarea.style.maxWidth = `${window.innerWidth - margin * 2}px`;
        textarea.style.zIndex = '11000';

        const measuredHeight = Math.max(textarea.getBoundingClientRect().height || 0, minHeight);
        const topIfBelow = triggerRect.bottom + gap;
        const topIfAbove = triggerRect.top - measuredHeight - gap;
        let top = placeBelow ? topIfBelow : topIfAbove;
        if (top < margin) top = margin;
        if (top + measuredHeight > window.innerHeight - margin) {
          top = Math.max(margin, window.innerHeight - measuredHeight - margin);
        }

        textarea.style.top = `${top}px`;

        // Reveal with animation after positioning
        textarea.style.opacity = '1';
        textarea.style.animation = 'popIn 0.18s cubic-bezier(0.175, 0.885, 0.32, 1.275) forwards';

        textarea.focus({ preventScroll: true });
      }
    }, 50);
  }

  closeReviewEdit() {
    if (this.editingReviewElement) {
      this.editingReviewElement.isEditingReview = false;
      this.editingReviewElement.isEditingHatyjaComments = false;
      this.editingReviewElement.isEditingEmailCommunications = false;
      this.editingReviewElement.isEditingRiskReasons = false;

      this.saveToStorage(this.editingReviewElement);
    }
    this.editingReviewElement = null;
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
      window: this.filterWindow,
      start: this.filterStartDate,
      end: this.filterEndDate,
      showArchived: this.showArchived,
    };
    const filterStr = JSON.stringify(filterValue);
    const filterChanged = filterStr !== this.lastAppliedFilter;
    this.lastAppliedFilter = filterStr;
    this.dataSource.filter = filterStr;

    if (filterChanged && this.dataSource.paginator) {
      this.dataSource.paginator.firstPage();
    }
  }

  onShowArchivedChange() {
    this.selection.clear();
    this.updateFilter();
    this.deferFilterCountRefresh();
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

  isAnyFilterActive(): boolean {
    return !!(
      (this.filterText && this.filterText.trim() !== '') ||
      this.filterStartDate ||
      this.filterEndDate ||
      (this.filterRegion && this.filterRegion !== 'All') ||
      (this.filterEntityType && this.filterEntityType !== 'All') ||
      (this.filterStatus && this.filterStatus !== 'All') ||
      (this.filterWindow && this.filterWindow !== 'All') ||
      this.showArchived
    );
  }

  clearAllFilters(searchInput?: HTMLInputElement) {
    if (searchInput) searchInput.value = '';
    this.filterText = '';
    this.filterStartDate = null;
    this.filterEndDate = null;
    this.filterRegion = 'All';
    this.filterEntityType = 'All';
    this.filterStatus = 'All';
    this.filterWindow = 'All';
    this.showArchived = false;
    this.updateFilter();
    this.deferFilterCountRefresh();
  }

  clearInvitedDate(element: ApplicantRecord, event: Event) {
    event.stopPropagation();
    element.invited = null;
    this.saveToStorage(element);
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

  onWindowFilterChange() {
    this.updateFilter();
  }

  private getRegionByCountry(country?: string): string {
    return getRegionByCountry(country || '', this.countryToRegionMap);
  }

  private matchesRegionSelection(rowRegion: string, selectedRegion?: string): boolean {
    if (!selectedRegion || selectedRegion === 'All') return true;
    return rowRegion === selectedRegion;
  }

  private getRegionCounts(): Record<string, number> {
    const rowRegions = this.getFilterCountSource().map((record) =>
      this.getRegionByCountry(record.country),
    );
    const counts: Record<string, number> = {};

    this.regionOptions.forEach((regionOption) => {
      counts[regionOption] = rowRegions.filter((rowRegion) =>
        this.matchesRegionSelection(rowRegion, regionOption),
      ).length;
    });

    return counts;
  }

  private refreshRegionOptionsWithCount() {
    const counts = this.getRegionCounts();
    this.regionOptionsWithCount = this.regionOptions.map((region) => ({
      value: region,
      label: `${region} (${counts[region] ?? 0})`,
    }));
  }

  private normalizeEntityType(entityType?: string): string {
    if (!entityType) return 'Unknown';
    const val = entityType.toLowerCase().trim();
    if (val.includes('un') && val.includes('system')) return 'UN System Entity';
    if (val.includes('un-system')) return 'UN System Entity';
    if (val.includes('government-controlled') || val.includes('government controlled'))
      return 'Public/Government-Controlled';
    if (val.includes('public') && val.includes('government') && val.includes('controlled'))
      return 'Public/Government-Controlled';
    if (val.includes('public') && val.includes('government')) return 'Public/Government';
    if (val === 'public/government' || val === 'public-government') return 'Public/Government';
    if (val.includes('private') || val.includes('sector')) return 'Private Sector';
    if (val.includes('multilateral') || val.includes('organization'))
      return 'Multilateral Organization';
    if (val.includes('non-profit') || val.includes('nonprofit')) return 'Multilateral Organization';
    if (val.includes('corporation') || val.includes('llc') || val.includes('partnership'))
      return 'Private Sector';
    return 'Unknown';
  }

  private matchesEntityTypeSelection(entityType?: string, selectedType?: string): boolean {
    if (!selectedType || selectedType === 'All') return true;
    return this.normalizeEntityType(entityType) === selectedType;
  }

  private getEntityTypeCounts(
    data: ApplicantRecord[],
    total: number,
  ): { name: string; count: number; pct: number }[] {
    const counts: Record<string, number> = {};
    data.forEach((d) => {
      const normalized = this.normalizeEntityType(d.entityType);
      counts[normalized] = (counts[normalized] || 0) + 1;
    });
    return Object.keys(counts)
      .map((name) => ({
        name,
        count: counts[name],
        pct: total > 0 ? Number(((counts[name] / total) * 100).toFixed(1)) : 0,
      }))
      .sort((a, b) => b.count - a.count);
  }

  private getFilterCountSource(): ApplicantRecord[] {
    return this.showArchived ? this.dataSource.data.filter((r) => r.archived) : this.activeRecords;
  }

  private getEntityTypeCounts2(): Record<string, number> {
    const counts: Record<string, number> = {};
    this.entityTypeOptions.forEach((opt) => {
      counts[opt] = 0;
    });
    this.getFilterCountSource().forEach((record) => {
      const normalized = this.normalizeEntityType(record.entityType);
      counts[normalized] = (counts[normalized] || 0) + 1;
      counts['All'] = (counts['All'] || 0) + 1;
    });
    return counts;
  }

  private refreshEntityTypeOptionsWithCount() {
    const counts = this.getEntityTypeCounts2();
    this.entityTypeOptionsWithCount = this.entityTypeOptions.map((opt) => ({
      value: opt,
      label: `${opt} (${counts[opt] ?? 0})`,
    }));
  }

  private refreshStatusOptionsWithCount() {
    const data = this.getFilterCountSource();
    const counts: Record<string, number> = { All: data.length };
    data.forEach((d) => {
      if (d.passed === 'Passed') counts['Passed'] = (counts['Passed'] || 0) + 1;
      else if (d.passed === 'Failed') counts['Failed'] = (counts['Failed'] || 0) + 1;
      else if (d.passed === 'Invited') counts['Invited'] = (counts['Invited'] || 0) + 1;
      else counts['Pending'] = (counts['Pending'] || 0) + 1;
    });
    this.statusOptionsWithCount = this.statusOptions.map((opt) => ({
      value: opt,
      label: `${opt} (${counts[opt] ?? 0})`,
    }));
  }

  private refreshWindowOptionsWithCount() {
    const data = this.getFilterCountSource();
    const counts: Record<string, number> = { All: data.length };
    data.forEach((d) => {
      const w = String(d.window ?? '');
      if (w) counts[w] = (counts[w] || 0) + 1;
    });
    this.windowOptionsWithCount = this.windowOptions.map((opt) => ({
      value: opt,
      label: `${opt} (${counts[opt] ?? 0})`,
    }));
  }

  private deferFilterCountRefresh() {
    this.refreshRegionOptionsWithCount();
    this.refreshEntityTypeOptionsWithCount();
    this.refreshStatusOptionsWithCount();
    this.refreshWindowOptionsWithCount();
  }

  // ── Toast Notifications ──────────────────────────────────────
  toast: { message: string; type: 'success' | 'error' | 'info'; visible: boolean } = {
    message: '',
    type: 'info',
    visible: false,
  };
  private toastTimer: any;

  showToast(message: string, type: 'success' | 'error' | 'info') {
    clearTimeout(this.toastTimer);
    this.toast = { message, type, visible: true };
    this.toastTimer = setTimeout(() => {
      this.toast.visible = false;
    }, 2500);
  }

  dismissToast() {
    this.toast.visible = false;
  }
  // ────────────────────────────────────────────────────────────────

  showClearConfirm = false;
  clearCountdown = 5;
  private clearCountdownTimer: any = null;

  openClearConfirm() {
    this.showClearConfirm = true;
    this.clearCountdown = 5;
    this.clearCountdownTimer = setInterval(() => {
      this.clearCountdown--;
      if (this.clearCountdown <= 0) {
        clearInterval(this.clearCountdownTimer);
        this.clearCountdownTimer = null;
      }
    }, 1000);
  }

  cancelClearConfirm() {
    this.showClearConfirm = false;
    clearInterval(this.clearCountdownTimer);
    this.clearCountdownTimer = null;
    this.clearCountdown = 5;
  }

  async clearData() {
    this.showClearConfirm = false;
    clearInterval(this.clearCountdownTimer);
    this.clearCountdownTimer = null;
    this.clearCountdown = 5;
    localStorage.removeItem(STORAGE_KEY);
    localStorage.removeItem(this.DIRTY_APPLICANTS_KEY);
    localStorage.removeItem(this.LOCAL_TIMESTAMP_KEY);
    this.dirtyApplicantIds.clear();
    await this.firestoreService.clearAllApplicants();
    this.dataSource.data = [];
    this.deferFilterCountRefresh();
    this.showToast('All data cleared.', 'info');
  }

  private firestoreSaveTimeout: any = null;

  saveToStorage(specificApplicant?: ApplicantRecord) {
    const data = this.dataSource.data;
    const now = new Date().toISOString();
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    localStorage.setItem(this.LOCAL_TIMESTAMP_KEY, now);

    if (specificApplicant?.id) {
      this.dirtyApplicantIds.set(specificApplicant.id, now);
      this.persistDirtyApplicants();
      this.scheduleFirestoreSave();
    }
  }

  private scheduleFirestoreSave() {
    clearTimeout(this.firestoreSaveTimeout);
    this.firestoreSaveTimeout = setTimeout(() => this.flushPendingLocalChanges(), 500);
  }

  private normalizeApplicantForSave(applicant: ApplicantRecord): any {
    const data: any = { ...applicant };
    // Convert submittedAt Date to ISO string for Firestore compatibility
    // Handle empty/undefined/null by setting to null so Firestore stores it correctly
    if (data.submittedAt instanceof Date) {
      data.submittedAt = data.submittedAt.toISOString();
    } else if (!data.submittedAt || String(data.submittedAt).trim() === '') {
      data.submittedAt = null; // Explicit null for Firestore
    }
    // invited date handling
    if (data.invited instanceof Date) {
      data.invited = data.invited.toISOString();
    } else if (!data.invited || String(data.invited).trim() === '') {
      data.invited = null;
    }
    // Remove UI-only flags before saving to Firestore
    delete data.isEditingReview;
    delete data.isEditingHatyjaComments;
    delete data.isEditingEmailCommunications;
    delete data.isEditingRiskReasons;
    delete data.isCalculatingRisk;
    return data;
  }

  private async pushApplicantToFirestore(
    applicant: ApplicantRecord,
  ): Promise<'saved' | 'remote_newer' | 'failed'> {
    if (!applicant.id) {
      applicant.id = this.generateApplicantId();
    }
    const id = applicant.id;
    const localDirtyAt = this.dirtyApplicantIds.get(id);

    try {
      const remote = await this.firestoreService.getApplicant(id);
      if (
        remote?.['updatedAt'] &&
        localDirtyAt &&
        this.isRemoteNewerThanLocal(remote['updatedAt'], localDirtyAt)
      ) {
        this.applyRemoteApplicantToLocal(id, remote);
        return 'remote_newer';
      }

      const { id: _, ...rest } = applicant;
      await this.firestoreService.saveApplicant(
        id,
        this.normalizeApplicantForSave(rest as ApplicantRecord),
      );
      this.dirtyApplicantIds.delete(id);
      this.persistDirtyApplicants();
      return 'saved';
    } catch (err) {
      console.error('Firestore: save applicant failed', err);
      return 'failed';
    }
  }

  private isRemoteNewerThanLocal(remoteUpdatedAt: string, localDirtyAt: string): boolean {
    return new Date(remoteUpdatedAt).getTime() > new Date(localDirtyAt).getTime();
  }

  private applyRemoteApplicantToLocal(id: string, remote: Record<string, any>) {
    const record = this.dataSource.data.find((r) => r.id === id);
    if (record) {
      Object.assign(record, { ...remote, id });
      this.hydrateRecords([record]);
    }
    this.dirtyApplicantIds.delete(id);
    this.persistDirtyApplicants();
    localStorage.setItem(STORAGE_KEY, JSON.stringify(this.dataSource.data));
  }

  private notifyRemoteVersionKept(count: number) {
    this.showToast(
      count === 1
        ? 'A newer Firebase version was kept; local change was not saved.'
        : `${count} records kept the newer Firebase version.`,
      'info',
    );
  }

  private async flushPendingLocalChanges(showReconnectToast = false) {
    if (this.dirtyApplicantIds.size === 0) return;

    const ids = [...this.dirtyApplicantIds.keys()];
    let synced = 0;
    let failed = 0;
    let remoteNewer = 0;

    for (const id of ids) {
      const record = this.dataSource.data.find((r) => r.id === id);
      if (!record) {
        this.dirtyApplicantIds.delete(id);
        continue;
      }
      const result = await this.pushApplicantToFirestore(record);
      if (result === 'saved') synced++;
      else if (result === 'remote_newer') remoteNewer++;
      else failed++;
    }

    if (synced > 0 || remoteNewer > 0) {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(this.dataSource.data));
    }
    if (synced > 0) {
      this.setFirestoreConnected(true);
      if (showReconnectToast) {
        this.showToast(
          synced === 1
            ? 'Offline change synced to Firebase.'
            : `${synced} offline changes synced to Firebase.`,
          'success',
        );
      }
    }
    if (remoteNewer > 0) {
      this.notifyRemoteVersionKept(remoteNewer);
    }
    if (failed > 0) {
      this.setFirestoreConnected(false, 'Failed to save changes to Firebase.');
    }
  }

  async saveApplicantImmediate(applicant: ApplicantRecord) {
    const now = new Date().toISOString();
    if (!applicant.id) {
      applicant.id = this.generateApplicantId();
    }
    if (!this.dirtyApplicantIds.has(applicant.id)) {
      this.dirtyApplicantIds.set(applicant.id, now);
      this.persistDirtyApplicants();
    }

    const result = await this.pushApplicantToFirestore(applicant);
    if (result === 'saved') {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(this.dataSource.data));
      this.setFirestoreConnected(true);
    } else if (result === 'remote_newer') {
      this.setFirestoreConnected(true);
      this.notifyRemoteVersionKept(1);
    } else {
      this.dirtyApplicantIds.set(applicant.id, now);
      this.persistDirtyApplicants();
      this.setFirestoreConnected(false, 'Failed to save changes to Firebase.');
    }
  }

  saveColumnVisibility() {
    localStorage.setItem('gfc_column_visibility', JSON.stringify(this.columnVisibility));
    this.firestoreService
      .saveColumnVisibility(this.columnVisibility)
      .then(() => this.setFirestoreConnected(true))
      .catch(() =>
        this.setFirestoreConnected(false, 'Failed to save column settings to Firebase.'),
      );
  }

  // Track the visibility of each column for selective export
  columnVisibility = {
    applicant: true,
    acronym: true,
    passed: true,
    drmcCompliance1: true,
    drmcCompliance2: true,
    entityType: true,
    submittedAt: true,
    preScreening: true,
    profiles: true,
    country: true,
    address: true,
    nolStatus: true,
    hatyjaReviewComments: true,
    redFlags: true,
    djResult: true,
    djReportNumber: true,
    djReportLink: true,
    djTruePositive: true,
    djFalsePositive: true,
    escalationRequired: true,
    hatyjaComments: true,
    emailCommunications: true,
    riskPercent: true,
    riskReasons: true,
    window: true,
    invited: true,
  };

  updateDisplayedColumns() {
    this.displayedColumns = [
      'select',
      ...this.columnKeys.filter((key) => (this.columnVisibility as any)[key]),
    ];
  }

  syncSelectedKeys() {
    this.selectedColumnKeys = this.columnKeys.filter((key) => (this.columnVisibility as any)[key]);
  }

  onColumnSelectionChange() {
    // Ensure `selectedColumnKeys` follow the canonical `columnKeys` order
    this.selectedColumnKeys = this.columnKeys.filter((k) => this.selectedColumnKeys.includes(k));

    // Reset all to false first
    Object.keys(this.columnVisibility).forEach((key) => {
      (this.columnVisibility as any)[key] = false;
    });
    // Set selected to true (in ordered sequence)
    this.selectedColumnKeys.forEach((key) => {
      (this.columnVisibility as any)[key] = true;
    });
    this.updateDisplayedColumns();
    this.saveColumnVisibility();
  }

  removeColumn(key: string) {
    this.selectedColumnKeys = this.selectedColumnKeys.filter((k) => k !== key);
    this.onColumnSelectionChange();
  }

  // Helper arrays for UI presentation
  columnKeys = [
    'applicant',
    'acronym',
    'passed',
    'window',
    'invited',
    'entityType',
    'submittedAt',
    'preScreening',
    'profiles',
    'country',
    'address',
    'nolStatus',
    'hatyjaReviewComments',
    'emailCommunications',
    'redFlags',
    'djResult',
    'djReportNumber',
    'djReportLink',
    'djTruePositive',
    'djFalsePositive',
    'escalationRequired',
    'hatyjaComments',
    'drmcCompliance1',
    'drmcCompliance2',
    'riskPercent',
    'riskReasons',
  ];
  columnNames: { [key: string]: string } = {
    applicant: 'Applicant',
    acronym: 'Acronym',
    passed: 'Status',
    drmcCompliance1: 'DRMC/Compliance 1',
    drmcCompliance2: 'DRMC/Compliance 2',
    entityType: 'Entity Type',
    submittedAt: 'Submitted At',
    preScreening: 'Pre-Screening',
    profiles: 'Profiles',
    country: 'Country',
    address: 'Address',
    nolStatus: 'NL Status',
    hatyjaReviewComments: 'Hatyja Review comments',
    redFlags: 'Red Flags',
    djResult: 'DJ-Result',
    djReportNumber: 'DJ report number',
    djReportLink: 'DJ report link',
    djTruePositive: 'DJ: True positive',
    djFalsePositive: 'DJ: False positive',
    escalationRequired: 'Escalation required to DRMC/Compliance?',
    hatyjaComments: 'Hatyja comments',
    emailCommunications: 'Email Communications',
    riskPercent: 'Risk %',
    riskReasons: 'Reasons',
    window: 'Window',
    invited: 'Invited',
  };

  // Selection helpers
  isAllSelected() {
    const visible = this.dataSource.filteredData;
    return visible.length > 0 && visible.every((row) => this.selection.isSelected(row));
  }

  setStatus(element: ApplicantRecord, status: string) {
    element.passed = status;
    if (status === 'Invited') {
      if (!element.invited) element.invited = new Date();
    }
    this.saveToStorage(element);
    this.deferFilterCountRefresh();
    this.updateFilter();
  }

  masterToggle() {
    const visible = this.dataSource.filteredData;
    if (this.isAllSelected()) {
      this.selection.clear();
    } else {
      visible.forEach((row) => this.selection.select(row));
    }
  }

  async archiveSelectedRows() {
    const selected = this.selection.selected;
    if (selected.length === 0) return;

    for (const row of selected) {
      row.archived = true;
      await this.saveApplicantImmediate(row);
    }
    this.selection.clear();
    this.deferFilterCountRefresh();
    this.updateFilter();
    this.showToast(`${selected.length} record(s) archived.`, 'info');
  }

  async unarchiveSelectedRows() {
    const selected = this.selection.selected;
    if (selected.length === 0) return;

    for (const row of selected) {
      row.archived = false;
      await this.saveApplicantImmediate(row);
    }
    this.selection.clear();
    this.deferFilterCountRefresh();
    this.updateFilter();
    this.showToast(`${selected.length} record(s) restored.`, 'info');
  }

  async deleteSelectedRows() {
    const selected = this.selection.selected;
    if (selected.length === 0) return;

    // Delete from Firestore individually
    for (const row of selected) {
      if (row.id) {
        await this.firestoreService.deleteApplicant(row.id);
      }
    }

    const data = this.dataSource.data;
    const newData = data.filter((row) => !this.selection.isSelected(row));
    this.dataSource.data = newData;
    this.deferFilterCountRefresh();
    this.selection.clear();
    localStorage.setItem(STORAGE_KEY, JSON.stringify(newData));
    this.showToast(`${selected.length} records deleted.`, 'info');
  }

  // Remove individual deleteRow as it's replaced by the selection model

  exportToExcel(selectedOnly: boolean): void {
    const rows = selectedOnly ? this.selection.selected : this.dataSource.filteredData;
    if (selectedOnly && rows.length === 0) {
      this.showToast('Select at least one row to export.', 'error');
      return;
    }

    const dataToExport = rows.map((row) => {
      const exportRow: Record<string, unknown> = {};
      this.columnKeys.forEach((key) => {
        if (!selectedOnly || (this.columnVisibility as any)[key]) {
          exportRow[this.columnNames[key]] = (row as any)[key];
        }
      });
      return exportRow;
    });

    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook: XLSX.WorkBook = { Sheets: { data: worksheet }, SheetNames: ['data'] };
    const suffix = selectedOnly ? 'Selected' : 'Export';
    XLSX.writeFile(workbook, `ApplicantData_${suffix}.xlsx`);
    this.showToast(
      selectedOnly ? `${rows.length} selected row(s) exported.` : 'Excel exported successfully!',
      'success',
    );
  }

  copyTableToClipboard(selectedOnly: boolean): void {
    const data = selectedOnly ? this.selection.selected : this.dataSource.filteredData;
    if (selectedOnly && data.length === 0) {
      this.showToast('Select at least one row to copy.', 'error');
      return;
    }
    const columns = this.columnKeys.filter((key) => (this.columnVisibility as any)[key]);

    // Create headers row
    const headers = columns.map((key) => this.columnNames[key]).join('\t');

    // Create data rowshatyja comments and reviews in blue
    const rows = data.map((row) => {
      return columns
        .map((key) => {
          let val = (row as any)[key];
          // Clean values for spreadsheet (remove newlines in reviews/addresses)
          if (typeof val === 'string') {
            val = val.replace(/\r?\n|\r/g, ' ');
          }
          return val || '';
        })
        .join('\t');
    });

    const tsv = [headers, ...rows].join('\n');

    navigator.clipboard
      .writeText(tsv)
      .then(() => {
        const msg = selectedOnly
          ? `${data.length} selected row(s) copied to clipboard.`
          : 'Copied to clipboard! You can now paste into Excel.';
        this.showToast(msg, 'success');
      })
      .catch((err) => {
        this.showToast('Failed to copy to clipboard.', 'error');
        console.error('Clipboard error:', err);
      });
  }

  async importExcel(event: any): Promise<void> {
    const target: DataTransfer = <DataTransfer>event.target;
    if (target.files.length !== 1) {
      this.showToast('Please select a single file.', 'error');
      return;
    }
    this.showToast('Importing data…', 'info');
    const reader: FileReader = new FileReader();
    reader.onload = async (e: any) => {
      this.importInProgress = true;
      const dataBuffer = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(dataBuffer, {
        type: 'array',
        cellDates: true,
        cellText: false,
        cellNF: true,
      });

      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      // Use header: 1 to get an array of arrays instead of objects. This avoids header name mismatch issues.
      const importedData: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });

      let existingData = [...this.dataSource.data];
      const existingMap = new Map<string, ApplicantRecord>();
      existingData.forEach((r) => {
        const key = (r.applicant || '').trim().toLowerCase();
        if (key) existingMap.set(key, r);
      });
      let addedCount = 0;
      let updatedCount = 0;

      if (importedData.length > 0) {
        // En lugar de asumir que la fila 0 es la cabecera, buscamos cuál fila contiene las cabeceras reales
        let headerRowIndex = 0;
        let idxApp = -1,
          idxAcronym = -1,
          idxType = -1,
          idxCountry = -1,
          idxAddress = -1,
          idxNol = -1,
          idxReview = -1,
          idxFlags = -1,
          idxPassed = -1;
        let idxSubmitted = -1,
          idxPreScreen = -1,
          idxProfiles = -1;
        let idxDjResult = -1,
          idxDjNumber = -1,
          idxDjLink = -1,
          idxDjTrue = -1,
          idxDjFalse = -1,
          idxEscalation = -1,
          idxHatyjaExtra = -1,
          idxEmailComm = -1;
        let idxDrmc1 = -1,
          idxDrmc2 = -1;
        let foundHeaders = false;

        for (let i = 0; i < importedData.length && i < 10; i++) {
          const row = importedData[i];
          if (!row || !row.length) continue;

          let matches = 0;
          let tempIdxApp = -1,
            tempIdxAcronym = -1,
            tempIdxType = -1,
            tempIdxCountry = -1,
            tempIdxAddress = -1,
            tempIdxNol = -1,
            tempIdxReview = -1,
            tempIdxFlags = -1,
            tempIdxPassed = -1;
          let tempIdxSub = -1,
            tempIdxPre = -1,
            tempIdxProf = -1;
          let tempIdxDjRes = -1,
            tempIdxDjNum = -1,
            tempIdxDjLnk = -1,
            tempIdxDjT = -1,
            tempIdxDjF = -1,
            tempIdxEsc = -1,
            tempIdxHatX = -1,
            tempIdxEmailComm = -1;
          let tempIdxDrmc1 = -1,
            tempIdxDrmc2 = -1;

          const headerCells = row.map((c: any, idx: number) => `[${idx}]="${c}"`).join(', ');
          console.log('Row', i, 'headers:', headerCells);

          row.forEach((col: any, index: number) => {
            if (!col) return;
            const colName = String(col).toLowerCase().trim();
            if (colName.includes('applicant') || colName === 'name' || colName === 'entity name') {
              tempIdxApp = index;
              matches++;
            } else if (colName.includes('acronym') || colName.includes('short name')) {
              tempIdxAcronym = index;
              matches++;
            } else if (colName.includes('entity') || colName.includes('type')) {
              tempIdxType = index;
              matches++;
            } else if (colName.includes('country') && !colName.includes('flag')) {
              tempIdxCountry = index;
              matches++;
            } else if (colName.includes('address') || colName.includes('location')) {
              tempIdxAddress = index;
              matches++;
            } else if (colName.includes('nol')) {
              tempIdxNol = index;
              matches++;
            } else if (
              colName === 'review' ||
              (colName.includes('hatyja') && colName.includes('review'))
            ) {
              tempIdxReview = index;
              matches++;
            } else if (
              colName.includes('red flag') ||
              colName.includes('red-flag') ||
              (colName.includes('flag') && !colName.includes('country'))
            ) {
              tempIdxFlags = index;
              matches++;
            } else if (colName.includes('passed') || colName.includes('status')) {
              tempIdxPassed = index;
              matches++;
            } else if (
              colName.includes('submit') ||
              colName.includes('subm') ||
              colName.includes('date') ||
              colName.includes('time') ||
              colName.includes('create') ||
              colName.includes('regist') ||
              colName.includes('enviado') ||
              colName.includes('fecha') ||
              colName === 'at'
            ) {
              tempIdxSub = index;
              matches++;
            } else if (colName.includes('screening') || colName.includes('pre-')) {
              tempIdxPre = index;
              matches++;
            } else if (colName.includes('dj-result') || colName.includes('dj result')) {
              tempIdxDjRes = index;
              matches++;
            } else if (colName.includes('report number') || colName.includes('dj report no')) {
              tempIdxDjNum = index;
              matches++;
            } else if (colName.includes('report link') || colName.includes('dj link')) {
              tempIdxDjLnk = index;
              matches++;
            } else if (
              colName.includes('profiles') ||
              colName.includes('profile') ||
              colName.includes('draft id') ||
              colName.includes('draft_id') ||
              colName.includes('url')
            ) {
              tempIdxProf = index;
              matches++;
            } else if (colName.includes('true positive')) {
              tempIdxDjT = index;
              matches++;
            } else if (colName.includes('false positive')) {
              tempIdxDjF = index;
              matches++;
            } else if (colName.includes('escalation')) {
              tempIdxEsc = index;
              matches++;
            } else if (colName.includes('email') && colName.includes('comm')) {
              tempIdxEmailComm = index;
              matches++;
            } else if (
              colName.includes('comments') &&
              (colName.includes('hatyja') || colName.includes('extra'))
            ) {
              tempIdxHatX = index;
              matches++;
            } else if (
              (colName.includes('drmc') || colName.includes('compliance')) &&
              (colName.includes('qa-1') ||
                colName.includes('qa 1') ||
                colName.includes('1') ||
                colName.includes('meixi'))
            ) {
              tempIdxDrmc1 = index;
              matches++;
            } else if (
              (colName.includes('drmc') || colName.includes('compliance')) &&
              (colName.includes('qa-2') || colName.includes('qa 2') || colName.includes('2'))
            ) {
              tempIdxDrmc2 = index;
              matches++;
            }
          });

          if (matches >= 2) {
            console.log('--- HEADER FOUND AT ROW', i, '---');
            console.log('Indices:', {
              applicant: tempIdxApp,
              acronym: tempIdxAcronym,
              type: tempIdxType,
              country: tempIdxCountry,
              submitted: tempIdxSub,
            });

            if (tempIdxSub === -1)
              console.warn('WARNING: Submitted column NOT found in this row headers.');

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
            idxEmailComm = tempIdxEmailComm;
            idxDrmc1 = tempIdxDrmc1;
            idxDrmc2 = tempIdxDrmc2;
            headerRowIndex = i;
            foundHeaders = true;
            console.log('Detected headers at row', i, { idxApp, idxCountry, idxSubmitted, idxNol });
            break;
          }
        }

        // Helper: extract a string value from a row cell
        const cellStr = (idx: number, row: any[]): string =>
          idx !== -1 && row[idx] !== undefined ? String(row[idx]).trim() : '';

        // Start reading data from the row AFTER the header
        for (let i = headerRowIndex + 1; i < importedData.length; i++) {
          const row = importedData[i];
          if (!row || !row.length) continue;
          const hasData = row.some(
            (cell: any) => cell !== undefined && cell !== null && String(cell).trim() !== '',
          );
          if (!hasData) continue;

          const rawDate = idxSubmitted !== -1 ? row[idxSubmitted] : undefined;
          if (i === headerRowIndex + 1) {
            console.log('Sample data row:', row);
            console.log('Raw date value found:', rawDate, typeof rawDate);
          }

          // Build the imported record from all detected columns
          const importedRow: Partial<ApplicantRecord> = {
            applicant: cellStr(idxApp, row),
            acronym: cellStr(idxAcronym, row),
            entityType: cellStr(idxType, row),
            country: cellStr(idxCountry, row),
            address: cellStr(idxAddress, row),
            nolStatus: cellStr(idxNol, row),
            hatyjaReviewComments: cellStr(idxReview, row),
            redFlags: cellStr(idxFlags, row),
            passed: (() => {
              const raw = cellStr(idxPassed, row);
              return this.validPassedValues.has(raw) ? raw : '';
            })(),
            submittedAt: this.parseImportedDate(rawDate),
            preScreening: cellStr(idxPreScreen, row),
            profiles: (() => {
              const val = cellStr(idxProfiles, row);
              return val && !val.startsWith('http')
                ? `https://partners.greenclimate.fund/pre-accreditation/${val}/staff/preview`
                : val;
            })(),
            djResult: this.normalizeDjDropdownImport('djResult', cellStr(idxDjResult, row)),
            djReportNumber: cellStr(idxDjNumber, row),
            djReportLink: cellStr(idxDjLink, row),
            djTruePositive: this.normalizeDjDropdownImport(
              'djTruePositive',
              cellStr(idxDjTrue, row),
            ),
            djFalsePositive: this.normalizeDjDropdownImport(
              'djFalsePositive',
              cellStr(idxDjFalse, row),
            ),
            escalationRequired: cellStr(idxEscalation, row),
            hatyjaComments: cellStr(idxHatyjaExtra, row),
            emailCommunications: cellStr(idxEmailComm, row),
            drmcCompliance1: cellStr(idxDrmc1, row),
            drmcCompliance2: cellStr(idxDrmc2, row),
          };

          const applicantKey = (importedRow.applicant || '').toLowerCase();
          const existing = applicantKey ? existingMap.get(applicantKey) : undefined;

          if (existing) {
            // Merge: only fill columns that are currently empty in the existing record
            let changed = false;
            const mergeFields: (keyof ApplicantRecord)[] = [
              'acronym',
              'entityType',
              'country',
              'address',
              'nolStatus',
              'hatyjaReviewComments',
              'emailCommunications',
              'redFlags',
              'passed',
              'submittedAt',
              'preScreening',
              'profiles',
              'djResult',
              'djReportNumber',
              'djReportLink',
              'djTruePositive',
              'djFalsePositive',
              'escalationRequired',
              'hatyjaComments',
              'drmcCompliance1',
              'drmcCompliance2',
            ];
            for (const field of mergeFields) {
              const existingVal = (existing as any)[field];
              const importedVal = (importedRow as any)[field];
              const isEmpty =
                existingVal === undefined ||
                existingVal === null ||
                String(existingVal).trim() === '';
              const hasImport =
                importedVal !== undefined &&
                importedVal !== null &&
                String(importedVal).trim() !== '';
              if (isEmpty && hasImport) {
                (existing as any)[field] = importedVal;
                changed = true;
              }
            }
            if (changed) updatedCount++;
          } else {
            // New applicant: create a full record
            const newRec: ApplicantRecord = {
              id: this.generateApplicantId(),
              applicant: importedRow.applicant || '',
              acronym: importedRow.acronym || '',
              entityType: importedRow.entityType || '',
              country: importedRow.country || '',
              address: importedRow.address || '',
              nolStatus: importedRow.nolStatus || '',
              hatyjaReviewComments: importedRow.hatyjaReviewComments || '',
              redFlags: importedRow.redFlags || '',
              passed: importedRow.passed || '',
              submittedAt: importedRow.submittedAt,
              preScreening: importedRow.preScreening || '',
              profiles: importedRow.profiles || '',
              djResult: importedRow.djResult || '',
              djReportNumber: importedRow.djReportNumber || '',
              djReportLink: importedRow.djReportLink || '',
              djTruePositive: importedRow.djTruePositive || '',
              djFalsePositive: importedRow.djFalsePositive || '',
              escalationRequired: importedRow.escalationRequired || '',
              hatyjaComments: importedRow.hatyjaComments || '',
              emailCommunications: importedRow.emailCommunications || '',
              drmcCompliance1: importedRow.drmcCompliance1 || '',
              drmcCompliance2: importedRow.drmcCompliance2 || '',
            };
            existingData.push(newRec);
            existingMap.set(applicantKey, newRec);
            addedCount++;
          }
        }
      }

      // Save all applicants to Firestore individually
      for (const applicant of existingData) {
        if (!applicant.id) applicant.id = this.generateApplicantId();
        await this.saveApplicantImmediate(applicant);
      }

      this.ngZone.run(() => {
        this.dataSource.data = existingData;
        this.deferFilterCountRefresh();
        localStorage.setItem(STORAGE_KEY, JSON.stringify(existingData));

        console.log('Final Records Sample:', existingData[0]);
        if (this.paginator) {
          this.paginator.firstPage();
        }
        const parts: string[] = [];
        if (addedCount > 0) parts.push(`${addedCount} new`);
        if (updatedCount > 0) parts.push(`${updatedCount} updated`);
        this.showToast(
          parts.length
            ? `Import done: ${parts.join(', ')} records.`
            : 'No changes — all data already up to date.',
          parts.length ? 'success' : 'info',
        );
        event.target.value = null;
        setTimeout(() => {
          this.importInProgress = false;
        }, 2000);
      });
    };
    reader.readAsArrayBuffer(target.files[0]);
  }

  async importExcel2(event: any): Promise<void> {
    const target: DataTransfer = <DataTransfer>event.target;
    if (target.files.length !== 1) {
      this.showToast('Please select a single file.', 'error');
      return;
    }
    this.showToast('Importing data (Excel 2)…', 'info');
    const reader: FileReader = new FileReader();
    reader.onload = async (e: any) => {
      this.importInProgress = true;
      const endImport = () => {
        setTimeout(() => {
          this.importInProgress = false;
        }, 2000);
      };
      const dataBuffer = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(dataBuffer, {
        type: 'array',
        cellDates: true,
        cellText: false,
        cellNF: true,
      });
      const ws: XLSX.WorkSheet = wb.Sheets[wb.SheetNames[0]];
      const raw: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });

      if (!raw.length) {
        this.showToast('Empty file.', 'error');
        event.target.value = null;
        endImport();
        return;
      }

      const nameToKey = new Map<string, string>();
      for (const [key, displayName] of Object.entries(this.columnNames)) {
        nameToKey.set(this.normalizeImportExcel2Header(displayName), key);
      }

      nameToKey.set('drmc/compliance qa-1 (meixi)', 'drmcCompliance1');
      nameToKey.set('drmc/compliance qa-1', 'drmcCompliance1');
      nameToKey.set('drmc/compliance qa 1', 'drmcCompliance1');
      nameToKey.set('drmc/compliance qa-2', 'drmcCompliance2');
      nameToKey.set('drmc/compliance qa 2', 'drmcCompliance2');
      nameToKey.set('compliance qa-1 (meixi)', 'drmcCompliance1');
      nameToKey.set('compliance qa-2', 'drmcCompliance2');
      nameToKey.set('hatyja comment', 'hatyjaComments');
      nameToKey.set('extra hatyja comments', 'hatyjaComments');

      // --- Detect header row (first 10 rows) and map column indices to field keys ---
      let headerRowIndex = -1;
      let colMap: { idx: number; key: string }[] = [];

      for (let r = 0; r < Math.min(raw.length, 10); r++) {
        const row = raw[r];
        if (!row || !row.length) continue;
        const tempMap: { idx: number; key: string }[] = [];
        row.forEach((cell: any, idx: number) => {
          const key = this.resolveImportExcel2ColumnKey(cell, nameToKey);
          if (!key) return;
          if (!tempMap.some((m) => m.key === key)) {
            tempMap.push({ idx, key });
          }
        });
        if (tempMap.length >= 2) {
          headerRowIndex = r;
          colMap = tempMap;
          console.log('Excel2: headers at row', r, colMap);
          break;
        }
      }

      if (headerRowIndex === -1) {
        this.showToast('Could not detect headers in file.', 'error');
        event.target.value = null;
        endImport();
        return;
      }

      const applicantMapping = colMap.find((m) => m.key === 'applicant');
      if (!applicantMapping) {
        this.showToast('No "Applicant" column found in file.', 'error');
        event.target.value = null;
        endImport();
        return;
      }

      // Build existing map
      const existingData = [...this.dataSource.data];
      const existingMap = new Map<string, ApplicantRecord>();
      existingData.forEach((rec) => {
        const k = (rec.applicant || '').trim().toLowerCase();
        if (k) existingMap.set(k, rec);
      });

      let matchCount = 0;
      let updatedCount = 0;
      let newCount = 0;
      let errorCount = 0;
      const applicantsToSave = new Set<ApplicantRecord>();

      const cellStr = (idx: number, row: any[]): string => this.importExcel2CellToString(row, idx);

      for (let i = headerRowIndex + 1; i < raw.length; i++) {
        const row = raw[i];
        if (!row || !row.length) continue;
        const hasData = row.some(
          (c: any) => c !== undefined && c !== null && String(c).trim() !== '',
        );
        if (!hasData) continue;

        try {
          const applicantName = cellStr(applicantMapping.idx, row);
          if (!applicantName) continue;

          const applicantKey = applicantName.trim().toLowerCase();
          const existing = existingMap.get(applicantKey);

          if (existing) {
            matchCount++;
            const fieldsUpdated = this.mergeImportExcel2IntoExisting(
              existing,
              colMap,
              row,
              cellStr,
            );
            if (fieldsUpdated.length > 0) {
              updatedCount++;
              applicantsToSave.add(existing);
              const now = new Date().toISOString();
              if (existing.id) {
                this.dirtyApplicantIds.set(existing.id, now);
              }
              console.log('Excel2 merged', applicantName, '→', fieldsUpdated.join(', '));
            }
          } else {
            const newRec: ApplicantRecord = {
              id: this.generateApplicantId(),
              applicant: applicantName,
            } as ApplicantRecord;
            for (const mapping of colMap) {
              if (mapping.key === 'applicant') continue;
              const val = this.importExcel2CellValue(mapping.key, mapping.idx, row, cellStr);
              const isDjDropdown =
                mapping.key === 'djResult' ||
                mapping.key === 'djTruePositive' ||
                mapping.key === 'djFalsePositive';
              if (isDjDropdown) {
                if (!this.isImportExcel2StringFieldEmpty(val)) {
                  (newRec as any)[mapping.key] = val;
                }
              } else if (val !== undefined && val !== null && String(val).trim() !== '') {
                (newRec as any)[mapping.key] = val;
              }
            }
            existingData.push(newRec);
            existingMap.set(applicantKey, newRec);
            applicantsToSave.add(newRec);
            newCount++;
          }
        } catch (err) {
          console.error('Excel2 import error at row', i, err);
          errorCount++;
        }
      }

      this.persistDirtyApplicants();

      for (const applicant of applicantsToSave) {
        if (!applicant.id) applicant.id = this.generateApplicantId();
        const now = new Date().toISOString();
        this.dirtyApplicantIds.set(applicant.id, now);
        this.persistDirtyApplicants();
        await this.pushApplicantToFirestore(applicant);
      }

      this.ngZone.run(() => {
        this.dataSource.data = [...existingData];
        this.deferFilterCountRefresh();
        localStorage.setItem(STORAGE_KEY, JSON.stringify(existingData));
        if (this.paginator) this.paginator.firstPage();

        const summary = [`${matchCount} matched`, `${updatedCount} updated`, `${newCount} new`];
        if (errorCount > 0) summary.push(`${errorCount} errors`);
        this.showToast(
          `Import 2 done: ${summary.join(', ')}.`,
          errorCount > 0 ? 'error' : 'success',
        );
        event.target.value = null;
        endImport();
      });
    };
    reader.readAsArrayBuffer(target.files[0]);
  }

  private normalizeImportExcel2Header(value: unknown): string {
    return String(value ?? '')
      .replace(/\u00a0/g, ' ')
      .toLowerCase()
      .trim()
      .replace(/\s+/g, ' ');
  }

  private resolveImportExcel2ColumnKey(
    cell: unknown,
    nameToKey: Map<string, string>,
  ): string | null {
    const norm = this.normalizeImportExcel2Header(cell);
    if (!norm) return null;

    if (nameToKey.has(norm)) return nameToKey.get(norm)!;

    // Mirror importExcel1 rules — review before generic "comments" matching
    if (norm === 'review' || (norm.includes('hatyja') && norm.includes('review'))) {
      return 'hatyjaReviewComments';
    }
    if (
      norm.includes('comments') &&
      (norm.includes('hatyja') || norm.includes('extra')) &&
      !norm.includes('review')
    ) {
      return 'hatyjaComments';
    }
    if (norm.includes('email') && norm.includes('comm')) return 'emailCommunications';
    if (
      norm.includes('red flag') ||
      norm.includes('red-flag') ||
      (norm.includes('flag') && !norm.includes('country'))
    ) {
      return 'redFlags';
    }
    if (norm.includes('passed') || norm === 'status') return 'passed';
    if (norm.includes('applicant') || norm === 'name' || norm === 'entity name') return 'applicant';
    if (norm.includes('acronym') || norm.includes('short name')) return 'acronym';
    if (norm.includes('entity') && norm.includes('type')) return 'entityType';
    if (norm.includes('country') && !norm.includes('flag')) return 'country';
    if (norm.includes('address') || norm.includes('location')) return 'address';
    if (norm.includes('nol')) return 'nolStatus';
    if (norm.includes('screening') || norm.includes('pre-')) return 'preScreening';
    if (
      norm.includes('profiles') ||
      norm.includes('profile') ||
      norm.includes('draft id') ||
      norm.includes('url')
    ) {
      return 'profiles';
    }
    if (norm.includes('dj-result') || norm.includes('dj result')) return 'djResult';
    if (norm.includes('report number') || norm.includes('dj report no')) return 'djReportNumber';
    if (norm.includes('report link') || norm.includes('dj link')) return 'djReportLink';
    if (norm.includes('true positive')) return 'djTruePositive';
    if (norm.includes('false positive')) return 'djFalsePositive';
    if (norm.includes('escalation')) return 'escalationRequired';
    if (
      (norm.includes('drmc') || norm.includes('compliance')) &&
      (norm.includes('qa-1') || norm.includes('qa 1') || norm.includes('meixi'))
    ) {
      return 'drmcCompliance1';
    }
    if (
      (norm.includes('drmc') || norm.includes('compliance')) &&
      (norm.includes('qa-2') || norm.includes('qa 2'))
    ) {
      return 'drmcCompliance2';
    }

    let best: { key: string; len: number } | null = null;
    for (const [dispNorm, key] of nameToKey.entries()) {
      if (norm.includes(dispNorm) || dispNorm.includes(norm)) {
        if (!best || dispNorm.length > best.len) {
          best = { key, len: dispNorm.length };
        }
      }
    }
    return best?.key ?? null;
  }

  private importExcel2CellToString(row: any[], idx: number): string {
    if (idx === -1 || row[idx] === undefined || row[idx] === null) return '';
    const cell = row[idx];
    if (cell instanceof Date) {
      return cell.toISOString();
    }
    if (typeof cell === 'object') {
      const richText = (cell as { text?: string; w?: string }).text ?? (cell as { w?: string }).w;
      if (richText !== undefined) return String(richText).trim();
    }
    return String(cell).trim();
  }

  private readonly importExcel2StringMergeFields = new Set<string>([
    'acronym',
    'entityType',
    'country',
    'address',
    'nolStatus',
    'hatyjaReviewComments',
    'redFlags',
    'passed',
    'preScreening',
    'profiles',
    'djResult',
    'djReportNumber',
    'djReportLink',
    'djTruePositive',
    'djFalsePositive',
    'escalationRequired',
    'hatyjaComments',
    'emailCommunications',
    'drmcCompliance1',
    'drmcCompliance2',
    'riskReasons',
  ]);

  private isImportExcel2StringFieldEmpty(value: unknown): boolean {
    if (value === undefined || value === null) return true;
    return String(value).trim() === '';
  }

  private readonly djResultDropdownOptions = ['Yes', 'No'] as const;
  private readonly djYesNoNaDropdownOptions = ['Yes', 'No', 'N/A'] as const;

  private normalizeDjDropdownImport(key: string, raw: string): string {
    const normalized = raw.trim().toLowerCase();
    if (!normalized) return '';

    if (key === 'djResult') {
      return this.djResultDropdownOptions.find((o) => o.toLowerCase() === normalized) ?? '';
    }
    if (key === 'djTruePositive' || key === 'djFalsePositive') {
      const aliases: Record<string, string> = {
        yes: 'Yes',
        no: 'No',
        'n/a': 'N/A',
        na: 'N/A',
        'n.a.': 'N/A',
      };
      if (aliases[normalized]) return aliases[normalized];
      return this.djYesNoNaDropdownOptions.find((o) => o.toLowerCase() === normalized) ?? '';
    }
    return '';
  }

  private isDjDropdownFieldEmpty(value: unknown): boolean {
    return this.isImportExcel2StringFieldEmpty(value);
  }

  private mergeImportExcel2IntoExisting(
    existing: ApplicantRecord,
    colMap: { idx: number; key: string }[],
    row: any[],
    cellStr: (idx: number, row: any[]) => string,
  ): string[] {
    const fieldsUpdated: string[] = [];
    for (const mapping of colMap) {
      if (mapping.key === 'applicant') continue;
      if (!this.importExcel2StringMergeFields.has(mapping.key)) continue;

      const existingVal = (existing as any)[mapping.key];
      const isDjDropdown =
        mapping.key === 'djResult' ||
        mapping.key === 'djTruePositive' ||
        mapping.key === 'djFalsePositive';
      const isEmpty = isDjDropdown
        ? this.isDjDropdownFieldEmpty(existingVal)
        : this.isImportExcel2StringFieldEmpty(existingVal);
      if (!isEmpty) continue;

      const importedVal = this.importExcel2CellValue(mapping.key, mapping.idx, row, cellStr);
      if (this.isImportExcel2StringFieldEmpty(importedVal)) continue;

      (existing as any)[mapping.key] = importedVal;
      fieldsUpdated.push(mapping.key);
    }
    return fieldsUpdated;
  }

  private importExcel2CellValue(
    key: string,
    idx: number,
    row: any[],
    cellStr: (idx: number, row: any[]) => string,
  ): any {
    if (key === 'submittedAt') {
      return this.parseImportedDate(row[idx]);
    }
    if (key === 'passed') {
      const raw = cellStr(idx, row);
      if (!raw) return '';
      const match = [...this.validPassedValues].find((v) => v.toLowerCase() === raw.toLowerCase());
      return match ?? '';
    }
    if (key === 'profiles') {
      const val = cellStr(idx, row);
      return val && !val.startsWith('http')
        ? `https://partners.greenclimate.fund/pre-accreditation/${val}/staff/preview`
        : val;
    }
    if (key === 'djResult' || key === 'djTruePositive' || key === 'djFalsePositive') {
      return this.normalizeDjDropdownImport(key, cellStr(idx, row));
    }
    return cellStr(idx, row);
  }

  // --- Investigation Skills Management ---
  private async loadSkills(): Promise<void> {
    this.investigationSkills = await this.skillsService.loadSkills();
  }

  saveSkills(): void {
    this.skillsService.saveSkills(this.investigationSkills);
    this.showToast('Investigation skills saved.', 'success');
  }

  addSkill(): void {
    this.skillsService.addSkill();
    this.investigationSkills = this.skillsService.getSkills();
  }

  removeSkill(index: number): void {
    const skill = this.investigationSkills[index];
    if (skill) {
      this.skillsService.removeSkill(skill.id);
      this.investigationSkills = this.skillsService.getSkills();
      this.saveSkills();
    }
  }

  resetSkills(): void {
    this.investigationSkills = this.skillsService.resetToDefaults();
    this.saveSkills();
  }

  // --- Risk Calculation ---
  promptForGeminiKey(): string | null {
    const key = prompt(
      'Enter your Gemini API Key.\nIt will be saved locally in your browser.',
      this.geminiApiKey || '',
    );
    if (key !== null && key.trim()) {
      this.geminiApiKey = key.trim();
      localStorage.setItem(this.GEMINI_KEY_STORAGE, this.geminiApiKey);
    }
    return key;
  }

  async calculateRisk(element: ApplicantRecord): Promise<void> {
    // Ensure API key is set
    if (!this.geminiApiKey) {
      const key = this.promptForGeminiKey();
      if (!key || !key.trim()) {
        this.showToast('Gemini API key is required.', 'error');
        return;
      }
    }

    // Update RiskService with current API key
    this.riskService.setApiKey(this.geminiApiKey);

    // Ask for corrections when recalculating (existing riskPercent)
    let userCorrection = '';
    if (element.riskPercent != null) {
      userCorrection =
        window.prompt(
          `Recalculating risk for "${element.applicant}".\n\nCurrent risk: ${element.riskPercent}%\n\n` +
            `If you want to help guide the analysis, enter corrections or additional context below:\n` +
            `(e.g., "website is actually example.org not example.com", "ignore results from 2015")\n\n` +
            `Leave empty to recalculate without changes:`,
          '',
        ) || '';
    }

    element.isCalculatingRisk = true;

    try {
      const result = await this.riskService.calculateRisk(
        element,
        this.investigationSkills,
        userCorrection || undefined,
      );

      this.ngZone.run(() => {
        element.riskPercent = result.riskPercent;
        element.riskReasons = result.riskReasons;
        element.isCalculatingRisk = false;
        this.saveToStorage(element);
        this.showToast(
          `Risk calculated for ${element.applicant}: ${element.riskPercent}%`,
          'success',
        );
      });
    } catch (err: any) {
      console.error('Risk calculation error:', err);

      // Handle specific error cases
      if (err.message?.includes('401') || err.message?.includes('403')) {
        this.geminiApiKey = '';
        localStorage.removeItem(this.GEMINI_KEY_STORAGE);
        this.showToast('Invalid API key. Please try again.', 'error');
      } else if (err.message?.includes('429')) {
        this.showToast('Rate limit exceeded. Please wait a moment and try again.', 'error');
      } else {
        this.showToast(`Error calculating risk: ${err.message || err}`, 'error');
      }

      this.ngZone.run(() => {
        element.isCalculatingRisk = false;
      });
    }
  }

  parseReasons(
    reasons: string | undefined,
  ): { text: string; type: 'positive' | 'negative' | 'neutral' }[] {
    if (!reasons) return [];
    return reasons
      .split('\n')
      .filter((l) => l.trim())
      .map((line) => {
        const trimmed = line.trim();
        if (trimmed.startsWith('[+]'))
          return { text: trimmed.substring(3).trim(), type: 'positive' as const };
        if (trimmed.startsWith('[-]'))
          return { text: trimmed.substring(3).trim(), type: 'negative' as const };
        return { text: trimmed, type: 'neutral' as const };
      });
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

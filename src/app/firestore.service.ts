import { Injectable } from '@angular/core';
import { db } from './firebase';
import {
  collection,
  doc,
  getDoc,
  getDocFromServer,
  setDoc,
  deleteDoc,
  getDocs,
  onSnapshot,
  query,
  Unsubscribe
} from 'firebase/firestore';

const APPLICANTS_COLLECTION = 'applicants';
const CONFIG_COLLECTION = 'gfc_data';

@Injectable({ providedIn: 'root' })
export class FirestoreService {

  // ── Individual Applicant Operations ─────────────────────────

  async saveApplicant(id: string, data: any): Promise<void> {
    try {
      await setDoc(doc(db, APPLICANTS_COLLECTION, id), {
        ...data,
        updatedAt: new Date().toISOString()
      });
    } catch (err) {
      console.error('Firestore: save applicant error', err);
      throw err;
    }
  }

  async getApplicant(id: string): Promise<Record<string, any> | null> {
    const ref = doc(db, APPLICANTS_COLLECTION, id);
    try {
      const snapshot = navigator.onLine ? await getDocFromServer(ref) : await getDoc(ref);
      return snapshot.exists() ? { id: snapshot.id, ...snapshot.data() } : null;
    } catch (err) {
      if (!navigator.onLine) {
        return null;
      }
      const snapshot = await getDoc(ref);
      return snapshot.exists() ? { id: snapshot.id, ...snapshot.data() } : null;
    }
  }

  async deleteApplicant(id: string): Promise<void> {
    try {
      await deleteDoc(doc(db, APPLICANTS_COLLECTION, id));
    } catch (err) {
      console.error('Firestore: delete applicant error', err);
      throw err;
    }
  }

  async loadAllApplicants(): Promise<any[]> {
    const q = query(collection(db, APPLICANTS_COLLECTION));
    const snapshot = await getDocs(q);
    return snapshot.docs.map(d => ({ id: d.id, ...d.data() }));
  }

  /**
   * Subscribe to real-time updates for all applicants
   * Returns unsubscribe function
   */
  subscribeToApplicants(
    onUpdate: (applicants: any[], fromServer: boolean) => void,
    onError?: (err: Error) => void
  ): Unsubscribe {
    const q = query(collection(db, APPLICANTS_COLLECTION));
    return onSnapshot(q, (snapshot) => {
      const applicants = snapshot.docs.map(d => ({ id: d.id, ...d.data() }));
      onUpdate(applicants, !snapshot.metadata.fromCache);
    }, (err) => {
      console.error('Firestore: subscription error', err);
      onError?.(err);
    });
  }

  // ── Migration from legacy format ────────────────────────────

  /**
   * Migrate data from old format (single document with array)
   * to new format (individual documents in applicants collection).
   * Returns number of records migrated.
   */
  async migrateFromLegacy(generateId: () => string): Promise<number> {
    try {
      // Check if already migrated (has data in new collection)
      const existing = await this.loadAllApplicants();
      if (existing.length > 0) {
        console.log('Migration: new collection already has data, skipping');
        return 0;
      }

      // Load legacy data
      const snapshot = await getDoc(doc(db, CONFIG_COLLECTION, 'applicant_records'));
      if (!snapshot.exists()) {
        console.log('Migration: no legacy data found');
        return 0;
      }

      const data = snapshot.data();
      const records = JSON.parse(data['records'] || '[]');

      if (!Array.isArray(records) || records.length === 0) {
        console.log('Migration: legacy data is empty');
        return 0;
      }

      console.log(`Migration: found ${records.length} records in legacy format`);

      // Migrate each record to individual document
      let migrated = 0;
      for (const record of records) {
        const id = record.id || generateId();
        const { id: _, ...dataWithoutId } = record;
        await this.saveApplicant(id, dataWithoutId);
        migrated++;
      }

      // Mark as migrated by renaming the old document (optional)
      await setDoc(doc(db, CONFIG_COLLECTION, 'applicant_records_migrated'), {
        originalUpdatedAt: data['updatedAt'],
        migratedAt: new Date().toISOString(),
        recordCount: migrated,
        migrated: true
      });

      console.log(`Migration: successfully migrated ${migrated} records`);
      return migrated;
    } catch (err) {
      console.error('Migration failed:', err);
      return 0;
    }
  }

  // ── Legacy batch operations (for migration/clear) ───────────

  async clearAllApplicants(): Promise<void> {
    try {
      const all = await this.loadAllApplicants();
      await Promise.all(all.map(a => this.deleteApplicant(a.id)));
    } catch (err) {
      console.error('Firestore: clear all error', err);
    }
  }

  // ── Investigation Skills ───────────────────────────────────

  async loadSkills(): Promise<any[] | null> {
    try {
      const snapshot = await getDoc(doc(db, CONFIG_COLLECTION, 'investigation_skills'));
      if (snapshot.exists()) {
        const skills = JSON.parse(snapshot.data()['skills'] || '[]');
        return Array.isArray(skills) && skills.length > 0 ? skills : null;
      }
    } catch (err) {
      console.warn('Firestore: load skills failed', err);
    }
    return null;
  }

  async saveSkills(skills: any[]): Promise<void> {
    try {
      await setDoc(doc(db, CONFIG_COLLECTION, 'investigation_skills'), {
        skills: JSON.stringify(skills),
        updatedAt: new Date().toISOString()
      });
    } catch (err) {
      console.error('Firestore: save skills error', err);
    }
  }

  // ── Column Visibility ──────────────────────────────────────

  async loadColumnVisibility(): Promise<Record<string, boolean> | null> {
    try {
      const snapshot = await getDoc(doc(db, CONFIG_COLLECTION, 'column_visibility'));
      if (snapshot.exists()) {
        return JSON.parse(snapshot.data()['data'] || '{}');
      }
    } catch (err) {
      console.warn('Firestore: load visibility failed', err);
    }
    return null;
  }

  async saveColumnVisibility(visibility: Record<string, boolean>): Promise<void> {
    try {
      await setDoc(doc(db, CONFIG_COLLECTION, 'column_visibility'), {
        data: JSON.stringify(visibility)
      });
    } catch (err) {
      console.error('Firestore: save visibility error', err);
    }
  }
}

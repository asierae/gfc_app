import { Injectable } from '@angular/core';
import { db } from './firebase';
import { doc, getDoc, setDoc } from 'firebase/firestore';

const COLLECTION = 'gfc_data';

@Injectable({ providedIn: 'root' })
export class FirestoreService {

  // ── Records ────────────────────────────────────────────────

  async loadRecords(): Promise<{ records: any[]; updatedAt: string } | null> {
    try {
      const snapshot = await getDoc(doc(db, COLLECTION, 'applicant_records'));
      if (snapshot.exists()) {
        const data = snapshot.data();
        const records = JSON.parse(data['records'] || '[]');
        if (Array.isArray(records)) {
          return { records, updatedAt: data['updatedAt'] || '' };
        }
      }
    } catch (err) {
      console.warn('Firestore: load records failed', err);
    }
    return null;
  }

  async saveRecords(records: any[]): Promise<void> {
    try {
      await setDoc(doc(db, COLLECTION, 'applicant_records'), {
        records: JSON.stringify(records),
        updatedAt: new Date().toISOString()
      });
      console.log(`Firestore: saved ${records.length} records`);
    } catch (err) {
      console.error('Firestore: save records error', err);
    }
  }

  async clearRecords(): Promise<void> {
    try {
      await setDoc(doc(db, COLLECTION, 'applicant_records'), {
        records: '[]',
        updatedAt: new Date().toISOString()
      });
    } catch (err) {
      console.error('Firestore: clear records error', err);
    }
  }

  // ── Investigation Skills ───────────────────────────────────

  async loadSkills(): Promise<any[] | null> {
    try {
      const snapshot = await getDoc(doc(db, COLLECTION, 'investigation_skills'));
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
      await setDoc(doc(db, COLLECTION, 'investigation_skills'), {
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
      const snapshot = await getDoc(doc(db, COLLECTION, 'column_visibility'));
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
      await setDoc(doc(db, COLLECTION, 'column_visibility'), {
        data: JSON.stringify(visibility)
      });
    } catch (err) {
      console.error('Firestore: save visibility error', err);
    }
  }
}

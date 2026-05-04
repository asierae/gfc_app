import { Injectable } from '@angular/core';
import { FirestoreService } from '../firestore.service';
import { InvestigationSkill, DEFAULT_SKILLS, SKILLS_STORAGE_KEY } from '../config/skills.config';

@Injectable({ providedIn: 'root' })
export class SkillsService {
  private skills: InvestigationSkill[] = [];

  constructor(private firestoreService: FirestoreService) {}

  /**
   * Load skills from localStorage and Firestore
   */
  async loadSkills(): Promise<InvestigationSkill[]> {
    // Load from localStorage first (immediate)
    const stored = localStorage.getItem(SKILLS_STORAGE_KEY);
    if (stored) {
      try {
        this.skills = JSON.parse(stored);
      } catch {
        this.skills = [...DEFAULT_SKILLS];
      }
    } else {
      this.skills = [...DEFAULT_SKILLS];
    }

    // Async load from Firestore
    const cloudSkills = await this.firestoreService.loadSkills();
    if (cloudSkills) {
      this.skills = cloudSkills;
      localStorage.setItem(SKILLS_STORAGE_KEY, JSON.stringify(cloudSkills));
    }

    return this.skills;
  }

  /**
   * Get current skills
   */
  getSkills(): InvestigationSkill[] {
    return this.skills;
  }

  /**
   * Save skills to localStorage and Firestore
   */
  async saveSkills(newSkills: InvestigationSkill[]): Promise<void> {
    this.skills = [...newSkills];
    localStorage.setItem(SKILLS_STORAGE_KEY, JSON.stringify(this.skills));
    await this.firestoreService.saveSkills(this.skills);
  }

  /**
   * Add a new skill
   */
  addSkill(label: string = 'New Skill', prompt: string = ''): InvestigationSkill {
    const newSkill: InvestigationSkill = {
      id: 'skill_' + Date.now(),
      label,
      prompt
    };
    this.skills.push(newSkill);
    return newSkill;
  }

  /**
   * Remove a skill by ID
   */
  removeSkill(id: string): void {
    this.skills = this.skills.filter(s => s.id !== id);
  }

  /**
   * Update a skill
   */
  updateSkill(id: string, updates: Partial<InvestigationSkill>): void {
    const skill = this.skills.find(s => s.id === id);
    if (skill) {
      Object.assign(skill, updates);
    }
  }

  /**
   * Reset to default skills
   */
  resetToDefaults(): InvestigationSkill[] {
    this.skills = [...DEFAULT_SKILLS];
    return this.skills;
  }

  /**
   * Generate skill prompts string for Gemini
   */
  generateSkillPrompts(): string {
    return this.skills.map(s => `### ${s.label}\n${s.prompt}`).join('\n\n');
  }
}

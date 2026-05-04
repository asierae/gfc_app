import { Injectable } from '@angular/core';
import { ApplicantRecord } from '../app.component';
import { InvestigationSkill, DEFAULT_SKILLS } from '../config/skills.config';
import { buildRiskAnalysisPrompt } from '../config/prompts.config';

interface GeminiResponse {
  candidates?: Array<{
    content?: {
      parts?: Array<{ text?: string }>;
    };
  }>;
}

export interface RiskCalculationResult {
  riskPercent: number;
  riskReasons: string;
}

@Injectable({ providedIn: 'root' })
export class RiskService {
  private geminiApiKey: string = '';

  setApiKey(key: string): void {
    this.geminiApiKey = key;
  }

  getApiKey(): string {
    return this.geminiApiKey;
  }

  /**
   * Calculate risk for an applicant using Gemini API
   */
  async calculateRisk(
    element: ApplicantRecord,
    skills: InvestigationSkill[],
    userCorrection?: string
  ): Promise<RiskCalculationResult> {
    if (!this.geminiApiKey) {
      throw new Error('Gemini API key is required');
    }

    const systemPrompt = buildRiskAnalysisPrompt({
      applicant: element.applicant,
      acronym: element.acronym,
      entityType: element.entityType,
      country: element.country,
      address: element.address,
      skills: skills.length > 0 ? skills : DEFAULT_SKILLS,
      userCorrection
    });

    const models = ['gemini-2.5-flash', 'gemini-2.5-pro'];
    const maxRetriesPerModel = 4;
    let lastError: Error | null = null;

    for (const model of models) {
      for (let attempt = 0; attempt <= maxRetriesPerModel; attempt++) {
        try {
          const response = await fetch(
            `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${this.geminiApiKey}`,
            {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                contents: [{ parts: [{ text: systemPrompt }] }],
                generationConfig: {
                  temperature: 0.2,
                  maxOutputTokens: 16384
                }
              })
            }
          );

          if ((response.status === 429 || response.status === 503) && attempt < maxRetriesPerModel) {
            const baseWait = response.status === 503 ? 5 : 2;
            const waitSec = baseWait * (attempt + 1);
            await new Promise(r => setTimeout(r, waitSec * 1000));
            continue;
          }

          if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
          }

          const data: GeminiResponse = await response.json();
          const text = data.candidates?.[0]?.content?.parts?.[0]?.text || '';

          if (!text) {
            throw new Error('Empty response from Gemini');
          }

          // Parse JSON from response (handle possible code fences)
          const jsonStr = text.replace(/```json\s*/g, '').replace(/```/g, '').trim();
          const parsed = JSON.parse(jsonStr);

          return {
            riskPercent: Math.max(0, Math.min(100, Math.round(parsed.riskPercent))),
            riskReasons: Array.isArray(parsed.reasons) ? parsed.reasons.join('\n') : String(parsed.reasons || '')
          };
        } catch (err) {
          lastError = err instanceof Error ? err : new Error(String(err));
          if (attempt < maxRetriesPerModel) {
            await new Promise(r => setTimeout(r, 1000 * (attempt + 1)));
          }
        }
      }
    }

    throw lastError || new Error('Failed to calculate risk after all retries');
  }

  /**
   * Parse imported date string to Date object
   */
  parseImportedDate(dateStr?: string): Date | string {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    return isNaN(d.getTime()) ? dateStr : d;
  }
}

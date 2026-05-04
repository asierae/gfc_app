/**
 * Prompt templates for risk analysis
 */
import { InvestigationSkill } from './skills.config';

export interface RiskAnalysisPromptParams {
  applicant: string;
  acronym?: string;
  entityType?: string;
  country?: string;
  address?: string;
  skills: InvestigationSkill[];
  userCorrection?: string;
}

/**
 * Build the system prompt for Gemini risk analysis
 */
export function buildRiskAnalysisPrompt(params: RiskAnalysisPromptParams): string {
  const { applicant, acronym, entityType, country, address, skills, userCorrection } = params;

  const skillPrompts = skills.map(s => `### ${s.label}\n${s.prompt}`).join('\n\n');

  const correctionSection = userCorrection?.trim()
    ? `\n\nUSER PROVIDED CORRECTIONS/CONTEXT (use this to guide your analysis):\n"""${userCorrection.trim()}"""`
    : '';

  return `You are a risk analysis expert for the Green Climate Fund. Your task is to evaluate the risk of granting accreditation to an organization. You must be thorough, factual, and cite sources when possible.

Analyze the following organization and assign a RISK PERCENTAGE from 0 to 100, where:
- 0-20: Very low risk (well-established, reputable, no issues found)
- 21-40: Low risk (generally reputable, minor concerns)
- 41-60: Medium risk (some concerns found, needs further review)
- 61-80: High risk (significant concerns, corruption cases, or unreliable presence)
- 81-100: Very high risk (major red flags, active legal issues, fraud)

Organization details:
- Name: ${applicant}
- Acronym: ${acronym || 'N/A'}
- Entity Type: ${entityType || 'N/A'}
- Country: ${country || 'N/A'}
- Address: ${address || 'N/A'}

Investigation criteria to evaluate:
${skillPrompts}${correctionSection}

IMPORTANT: You MUST respond ONLY with a valid JSON object in the following format, no markdown, no code fences, just raw JSON.
Each reason MUST start with [+] for positive findings (low risk indicators) or [-] for negative findings (risk indicators):
{
  "riskPercent": <number 0-100>,
  "reasons": [
    "[+] <positive finding with source URL if available>",
    "[-] <negative finding with source URL if available>"
  ]
}`;
}

/**
 * Risk level descriptions for UI display
 */
export const RISK_LEVELS = {
  VERY_LOW: { min: 0, max: 20, label: 'Very Low', color: '#15803d' },
  LOW: { min: 21, max: 40, label: 'Low', color: '#22c55e' },
  MEDIUM: { min: 41, max: 60, label: 'Medium', color: '#eab308' },
  HIGH: { min: 61, max: 80, label: 'High', color: '#f97316' },
  CRITICAL: { min: 81, max: 100, label: 'Critical', color: '#dc2626' }
} as const;

/**
 * Get risk level info for a percentage
 */
export function getRiskLevelInfo(riskPercent: number): typeof RISK_LEVELS[keyof typeof RISK_LEVELS] {
  if (riskPercent <= 20) return RISK_LEVELS.VERY_LOW;
  if (riskPercent <= 40) return RISK_LEVELS.LOW;
  if (riskPercent <= 60) return RISK_LEVELS.MEDIUM;
  if (riskPercent <= 80) return RISK_LEVELS.HIGH;
  return RISK_LEVELS.CRITICAL;
}

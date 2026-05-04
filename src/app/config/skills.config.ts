/**
 * Default investigation skills for risk analysis
 */
export interface InvestigationSkill {
  id: string;
  label: string;
  prompt: string;
}

export const DEFAULT_SKILLS: InvestigationSkill[] = [
  {
    id: 'webpage_check',
    label: 'Website Reliability',
    prompt: 'Check if the organization has an official website. Verify the domain registration date, SSL certificate, and professional appearance. Look for contact information, about us sections, and transparency. Flag if no website is found or if it looks suspicious/placeholder.'
  },
  {
    id: 'linkedin_check',
    label: 'LinkedIn Presence',
    prompt: 'Search for the organization on LinkedIn. Check if they have an official company page with active employees, company size, and recent activity. Look for the leadership team profiles. Flag if no LinkedIn presence or if it appears minimal/inactive.'
  },
  {
    id: 'facebook_check',
    label: 'Facebook Activity',
    prompt: 'Search for the organization on Facebook. Check for an official page, follower count, post frequency, and community engagement. Look for reviews and comments. Flag if no page exists or if there are negative reviews/scam reports.'
  },
  {
    id: 'google_news',
    label: 'News Coverage',
    prompt: 'Search Google News for mentions of the organization. Look for press releases, media coverage, partnerships, and any controversies. Positive coverage reduces risk; negative news or no coverage increases risk.'
  },
  {
    id: 'corruption_check',
    label: 'Corruption & Scandals',
    prompt: 'Search for corruption cases, scandals, fraud allegations, or legal issues involving the organization or its key leadership. Check court records, investigative journalism, and anti-corruption databases. Major red flags should significantly increase risk score.'
  },
  {
    id: 'ngo_check',
    label: 'NGO Database Registration',
    prompt: 'Check if the organization is registered in major NGO databases like OpenAID, Devex, or national NGO registries. Verify their registration status, transparency ratings, and reported activities.'
  },
  {
    id: 'gov_registration',
    label: 'Government Registration',
    prompt: 'Verify government registration in their home country. Check business registries, tax identification numbers, and official government databases. Unregistered or suspicious registrations are major red flags.'
  }
];

export const SKILLS_STORAGE_KEY = 'gfc_investigation_skills';

/**
 * Country and region configuration for geographic filtering
 */

export const REGION_OPTIONS = [
  'All',
  'APAC',
  'LAC',
  'AFR',
  'ECM',
  'INTERNATIONAL'
] as const;

export type RegionOption = typeof REGION_OPTIONS[number];

export const REGION_COUNTRY_MAP: Record<string, Set<string>> = {
  APAC: new Set([
    'afghanistan', 'bangladesh', 'bhutan', 'china', 'hong kong', 'india', 'japan',
    'macau', 'maldives', 'mongolia', 'nepal', 'north korea', 'pakistan',
    'philippines', 'singapore', 'south korea', 'sri lanka', 'thailand', 'timor-leste',
    'vietnam', 'Viet Nam', 'brunei', 'cambodia', 'indonesia', 'malaysia', 'myanmar',
    'papua new guinea', 'samoa', 'solomon islands', 'tonga', 'tuvalu', 'vanuatu', 'Lao People\'s Democratic Republic (the)'
  ]),
  LAC: new Set([
    'argentina', 'bahamas', 'barbados', 'belize', 'bolivia', 'brazil', 'chile', 'colombia', 'costa rica',
    'cuba', 'dominican republic', 'ecuador', 'el salvador', 'guatemala', 'guyana', 'haiti', 'honduras',
    'jamaica', 'mexico', 'nicaragua', 'panama', 'paraguay', 'peru', 'uruguay', 'venezuela'
  ]),
  AFR: new Set([
    'algeria', 'angola', 'benin', 'botswana', 'burkina faso', 'burundi', 'cameroon', 'cape verde', 'côte d\'ivoire', 'Central African Republic (the)',
    'central african republic', 'chad', 'comoros', 'congo', 'democratic republic of the congo',
    'djibouti', 'egypt', 'equatorial guinea', 'eritrea', 'eswatini', 'ethiopia', 'gabon', 'gambia',
    'ghana', 'guinea', 'guinea-bissau', 'ivory coast', 'kenya', 'lesotho', 'liberia', 'libya',
    'madagascar', 'malawi', 'mali', 'mauritania', 'mauritius', 'morocco', 'mozambique', 'namibia',
    'niger', 'nigeria', 'rwanda', 'sao tome and principe', 'senegal', 'seychelles', 'sierra leone',
    'somalia', 'south africa', 'south sudan', 'sudan', 'tanzania', 'togo', 'tunisia', 'uganda',
    'zambia', 'zimbabwe'
  ]),
  ECM: new Set([
    'albania', 'armenia', 'azerbaijan', 'belarus', 'bosnia and herzegovina', 'bulgaria', 'croatia',
    'cyprus', 'czech republic', 'estonia', 'georgia', 'hungary', 'latvia', 'lithuania', 'moldova',
    'montenegro', 'north macedonia', 'poland', 'romania', 'russia', 'serbia', 'slovakia', 'slovenia',
    'turkey', 'ukraine', 'austria', 'belgium', 'denmark', 'finland', 'france', 'germany', 'greece',
    'iceland', 'ireland', 'italy', 'luxembourg', 'netherlands', 'norway', 'portugal', 'spain',
    'sweden', 'switzerland', 'united kingdom',
    'bahrain', 'iran', 'iraq', 'israel', 'jordan', 'kuwait', 'lebanon', 'oman', 'palestine', 'qatar',
    'saudi arabia', 'syria', 'united arab emirates', 'yemen', 'kyrgyzstan', 'tajikistan', 'turkmenistan', 'uzbekistan'
  ]),
  INTERNATIONAL: new Set([
    'united nations', 'world bank', 'international', 'global', 'multilateral'
  ])
};

/**
 * Build a reverse map from country name to region
 */
export function buildCountryToRegionMap(): Map<string, string> {
  const map = new Map<string, string>();
  Object.entries(REGION_COUNTRY_MAP).forEach(([region, countries]) => {
    countries.forEach(country => {
      map.set(normalizeCountryName(country), region);
    });
  });
  return map;
}

/**
 * Normalize country name for comparison
 */
export function normalizeCountryName(country?: string): string {
  if (!country) return '';
  return country
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Get region for a given country
 */
export function getRegionByCountry(country: string, countryToRegionMap: Map<string, string>): string {
  const normalized = normalizeCountryName(country);
  if (!normalized) return 'INTERNATIONAL';
  return countryToRegionMap.get(normalized) || 'INTERNATIONAL';
}

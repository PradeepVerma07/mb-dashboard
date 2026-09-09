/**
 * siteHierarchy.js (client) – Atomic Physical Panel Conflict Map
 *
 * Mirror of server/src/siteHierarchy.js for use in the frontend.
 * Used to show live conflict warnings when selecting sites in
 * Campaign modal, Proposal builder, and PPT generator.
 */

export const SITE_PANELS = {
  // ── Group 1: Shivranjani Bridge – Nr. D-Mart Junction
  'MB-01': ['SHV_DM_L'],
  'MB-02': ['SHV_DM_M'],
  'MB-03': ['SHV_DM_R'],
  'MB-04': ['SHV_DM_L','SHV_DM_M','SHV_DM_R'],
  'MB-05': ['SHV_DM_M','SHV_DM_L'],
  'MB-06': ['SHV_DM_M','SHV_DM_R'],

  // ── Group 2: Shivranjani Bridge – Fcg. Shyamal
  'MB-07': ['SHV_SH_L'],
  'MB-08': ['SHV_SH_M'],
  'MB-09': ['SHV_SH_R'],
  'MB-10': ['SHV_SH_L','SHV_SH_M','SHV_SH_R'],
  'MB-11': ['SHV_SH_M','SHV_SH_L'],
  'MB-12': ['SHV_SH_M','SHV_SH_R'],

  // ── Group 3: Sindhubhavan Road
  'MB-17': ['SINDH_U','SINDH_L'],
  'MB-18': ['SINDH_U'],
  'MB-19': ['SINDH_L'],

  // ── Group 4: Zaveri Circle
  'MB-22': ['ZAV_L','ZAV_R'],
  'MB-23': ['ZAV_L'],
  'MB-24': ['ZAV_R'],

  // ── Group 5: Thaltej PVR
  'MB-31': ['PVR_U','PVR_L'],
  'MB-32': ['PVR_U'],
  'MB-33': ['PVR_L'],

  // ── Group 6: Acropolis Mall
  'MB-34': ['ACRO_L','ACRO_R'],
  'MB-35': ['ACRO_L'],
  'MB-36': ['ACRO_R'],

  // ── Group 7: Hebatpur – Zydus
  'MB-38': ['HEB_U','HEB_L'],
  'MB-39': ['HEB_U'],
  'MB-40': ['HEB_L'],

  // ── Group 8: Adani Shantigram
  'MB-41': ['SHAN_L','SHAN_R'],
  'MB-42': ['SHAN_L'],
  'MB-43': ['SHAN_R'],

  // ── Group 9: Keshavbaug / ITC Narmada
  'MB-44': ['KESH_U','KESH_L'],
  'MB-45': ['KESH_U'],
  'MB-46': ['KESH_L'],

  // ── Group 10: Jodhpur Cross Road – Satellite
  'MB-47': ['JCR_U','JCR_L'],
  'MB-48': ['JCR_U'],
  'MB-49': ['JCR_L'],

  // ── Group 11: Prahladnagar – Towards Anandnagar
  'MB-50': ['PRN_AN_U','PRN_AN_L'],
  'MB-51': ['PRN_AN_U'],
  'MB-52': ['PRN_AN_L'],

  // ── Group 12: Prahladnagar – Towards Shyamal Cross
  'MB-53': ['PRN_SH_U','PRN_SH_L'],
  'MB-54': ['PRN_SH_U'],
  'MB-55': ['PRN_SH_L'],

  // ── Group 13: Prahladnagar – Shyamal to Anandnagar
  'MB-56': ['PRN_SA_U','PRN_SA_L'],
  'MB-57': ['PRN_SA_U'],
  'MB-58': ['PRN_SA_L'],

  // ── Group 14: Sun-N-Step Club
  'MB-60': ['SNS_U','SNS_L'],
  'MB-61': ['SNS_U'],
  'MB-62': ['SNS_L'],

  // ── Group 15: C.G Road – Wagh Bakri
  'MB-65': ['CG_U','CG_L'],
  'MB-66': ['CG_U'],
  'MB-67': ['CG_L'],

  // ── Group 16: South Bopal – Sobo Center
  'MB-68': ['SOBO_U','SOBO_L'],
  'MB-69': ['SOBO_U'],
  'MB-70': ['SOBO_L'],

  // ── Group 17: 200ft Ring Road – Left
  'MB-71': ['RRL_U','RRL_L'],
  'MB-72': ['RRL_U'],
  'MB-73': ['RRL_L'],

  // ── Group 18: 200ft Ring Road – Right
  'MB-74': ['RRR_U','RRR_L'],
  'MB-75': ['RRR_U'],
  'MB-76': ['RRR_L'],

  // ── Group 19: Billionaire Street
  'MB-77': ['BILL_U','BILL_L'],
  'MB-78': ['BILL_U'],
  'MB-79': ['BILL_L'],

  // ── Group 20: Bopal Road – TRP Mall
  'MB-80': ['TRP_U','TRP_L'],
  'MB-81': ['TRP_U'],
  'MB-82': ['TRP_L'],

  // ── Group 21: Sanathal Circle Bridge
  'MB-85': ['SAT_U','SAT_L'],
  'MB-86': ['SAT_U'],
  'MB-87': ['SAT_L'],
};

const PANEL_TO_SITES = {};
for (const [code, panels] of Object.entries(SITE_PANELS)) {
  for (const panel of panels) {
    if (!PANEL_TO_SITES[panel]) PANEL_TO_SITES[panel] = [];
    PANEL_TO_SITES[panel].push(code);
  }
}

/** Returns all site codes that share any physical panel with siteCode (includes itself). */
export function getOverlappingSiteCodes(siteCode) {
  const code = String(siteCode || '').toUpperCase().trim();
  const myPanels = SITE_PANELS[code];
  if (!myPanels || myPanels.length === 0) return [code];
  const affected = new Set();
  for (const panel of myPanels) {
    for (const s of (PANEL_TO_SITES[panel] || [])) affected.add(s);
  }
  return Array.from(affected);
}

/**
 * Given the full sites array from the server (each with .site_code & .availability),
 * and a new siteCode being selected for a campaign,
 * returns the list of site_codes that will also become occupied.
 */
export function getLinkedWillBeOccupied(siteCode) {
  const code = String(siteCode || '').toUpperCase().trim();
  return getOverlappingSiteCodes(code).filter(c => c !== code);
}

/**
 * Given a list of already-occupied site codes (from active campaigns),
 * returns a Map<siteCode, reason> covering direct + linked sites.
 */
export function computeAllOccupied(bookedCodes) {
  const result = new Map();
  for (const code of bookedCodes) {
    const upper = String(code || '').toUpperCase().trim();
    if (!upper) continue;
    if (!result.has(upper)) result.set(upper, 'Direct');
    for (const linked of getOverlappingSiteCodes(upper)) {
      if (linked !== upper && !result.has(linked)) {
        result.set(linked, `Linked via ${upper}`);
      }
    }
  }
  return result;
}

/** Canonicalize site code to standard 'MB-XX' format */
export function canonicalSiteCode(code) {
  if (!code) return '';
  const clean = String(code).trim().toUpperCase();
  const m = clean.match(/^MB[-_ ]?(\d+)$/i);
  if (m) {
    return `MB-${m[1].padStart(2, '0')}`;
  }
  return clean;
}

/** Check if site is a combined/multi-panel structure */
export function isCombinedSite(siteCode) {
  const code = canonicalSiteCode(siteCode);
  const panels = SITE_PANELS[code];
  return Boolean(panels && panels.length > 1);
}

/** Return classification: 'Combined' | 'Split Face' | 'Single' */
export function getSiteTypeTag(siteCode) {
  const code = canonicalSiteCode(siteCode);
  const panels = SITE_PANELS[code];
  if (!panels) return 'Single';
  if (panels.length > 1) return 'Combined';
  const overlapping = getOverlappingSiteCodes(code);
  return overlapping.length > 1 ? 'Split Face' : 'Single';
}

/** Return structured conflict summary for UI warning */
export function getConflictSummary(siteCode) {
  const code = canonicalSiteCode(siteCode);
  const overlapping = getOverlappingSiteCodes(code).filter(c => c !== code);
  if (overlapping.length === 0) return null;
  const isComb = isCombinedSite(code);
  return {
    isCombined: isComb,
    type: isComb ? 'Combined Site' : 'Split Face',
    overlapping,
    message: isComb
      ? `Combined site: Booking ${code} automatically occupies linked sites: ${overlapping.join(', ')}`
      : `Split face: Booking ${code} automatically blocks combined sites: ${overlapping.join(', ')}`
  };
}

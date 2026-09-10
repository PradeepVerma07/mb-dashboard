/**
 * siteHierarchy.js – Atomic Physical Panel Conflict Map
 *
 * Every site is expressed as a set of atomic physical panels.
 * If two sites share ANY panel, they cannot be simultaneously booked.
 *
 * Booking one site automatically marks all overlapping sites as Occupied.
 * When the campaign ends / is deleted, the sync recomputes and restores availability.
 */

// ─── Atomic Panel Definitions ────────────────────────────────────────────────
// Each key is a site_code; value is an array of atomic panel IDs it occupies.

export const SITE_PANELS = {
  // ── Group 1: Shivranjani Bridge – Nr. D-Mart Junction ─────────────────────
  // Panels: SHV_DM_L (Left), SHV_DM_M (Middle), SHV_DM_R (Right)
  'MB-01': ['SHV_DM_L'],                      // Left only
  'MB-02': ['SHV_DM_M'],                      // Middle only
  'MB-03': ['SHV_DM_R'],                      // Right only
  'MB-04': ['SHV_DM_L','SHV_DM_M','SHV_DM_R'], // Full 1-3
  'MB-05': ['SHV_DM_M','SHV_DM_L'],           // Middle+Left (2-1)
  'MB-06': ['SHV_DM_M','SHV_DM_R'],           // Middle+Right (2-3)

  // ── Group 2: Shivranjani Bridge – Fcg. Shyamal 4 Road Jnc ────────────────
  // Panels: SHV_SH_L, SHV_SH_M, SHV_SH_R
  'MB-07': ['SHV_SH_L'],
  'MB-08': ['SHV_SH_M'],
  'MB-09': ['SHV_SH_R'],
  'MB-10': ['SHV_SH_L','SHV_SH_M','SHV_SH_R'], // Full 1-3
  'MB-11': ['SHV_SH_M','SHV_SH_L'],             // Middle+Left (2-1)
  'MB-12': ['SHV_SH_M','SHV_SH_R'],             // Middle+Right (2-3)

  // ── Group 3: Sindhubhavan Road ─────────────────────────────────────────────
  // Panels: SINDH_U (Upper), SINDH_L (Lower)
  'MB-17': ['SINDH_U','SINDH_L'],   // Full 30x31
  'MB-18': ['SINDH_U'],             // Upper
  'MB-19': ['SINDH_L'],             // Lower

  // ── Group 4: Zaveri Circle – Bopal Ambli Road ─────────────────────────────
  // Panels: ZAV_L, ZAV_R
  'MB-22': ['ZAV_L','ZAV_R'],       // Full 31x15
  'MB-23': ['ZAV_L'],               // Left
  'MB-24': ['ZAV_R'],               // Right

  // ── Group 5: Thaltej Junction PVR – S.G. Highway ─────────────────────────
  // Panels: PVR_U, PVR_L
  'MB-31': ['PVR_U','PVR_L'],       // Full 30x31
  'MB-32': ['PVR_U'],               // Upper
  'MB-33': ['PVR_L'],               // Lower

  // ── Group 6: Thaltej Acropolis Mall – S.G. Highway ───────────────────────
  // Panels: ACRO_L, ACRO_R
  'MB-34': ['ACRO_L','ACRO_R'],     // Full 42x20
  'MB-35': ['ACRO_L'],              // Left
  'MB-36': ['ACRO_R'],              // Right

  // ── Group 7: Hebatpur Road – Bagaban Circle to Zydus ─────────────────────
  // Panels: HEB_U, HEB_L
  'MB-38': ['HEB_U','HEB_L'],       // Full 30x21
  'MB-39': ['HEB_U'],               // Upper
  'MB-40': ['HEB_L'],               // Lower

  // ── Group 8: Adani Shantigram – Khoraj Flyover ───────────────────────────
  // Panels: SHAN_L, SHAN_R
  'MB-41': ['SHAN_L','SHAN_R'],     // Full 80x20
  'MB-42': ['SHAN_L'],              // Left
  'MB-43': ['SHAN_R'],              // Right

  // ── Group 9: Keshavbaug / ITC Narmada – Judges Bungalow ─────────────────
  // Panels: KESH_U, KESH_L
  'MB-44': ['KESH_U','KESH_L'],     // Full 20x31
  'MB-45': ['KESH_U'],              // Upper
  'MB-46': ['KESH_L'],              // Lower

  // ── Group 10: Jodhpur Cross Road – Satellite ──────────────────────────────
  // Panels: JCR_U, JCR_L
  'MB-47': ['JCR_U','JCR_L'],       // Full 15x31
  'MB-48': ['JCR_U'],               // Upper
  'MB-49': ['JCR_L'],               // Lower

  // ── Group 11: Prahladnagar Road – Towards Anandnagar ─────────────────────
  // Panels: PRN_AN_U, PRN_AN_L
  'MB-50': ['PRN_AN_U','PRN_AN_L'], // Full 30x31
  'MB-51': ['PRN_AN_U'],            // Upper (Sachin Tower)
  'MB-52': ['PRN_AN_L'],            // Lower

  // ── Group 12: Prahladnagar Road – Towards Shyamal Cross ──────────────────
  // Panels: PRN_SH_U, PRN_SH_L
  'MB-53': ['PRN_SH_U','PRN_SH_L'], // Full 15x31
  'MB-54': ['PRN_SH_U'],            // Upper
  'MB-55': ['PRN_SH_L'],            // Lower

  // ── Group 13: Prahladnagar Road – Shyamal to Anandnagar ─────────────────
  // Panels: PRN_SA_U, PRN_SA_L
  'MB-56': ['PRN_SA_U','PRN_SA_L'], // Full 15x31
  'MB-57': ['PRN_SA_U'],            // Upper
  'MB-58': ['PRN_SA_L'],            // Lower

  // ── Group 14: Sun-N-Step Club – Judges Bungalow ──────────────────────────
  // Panels: SNS_U, SNS_L
  'MB-60': ['SNS_U','SNS_L'],       // Full 30x31
  'MB-61': ['SNS_U'],               // Upper
  'MB-62': ['SNS_L'],               // Lower

  // ── Group 15: Commerce College Road – C.G Road ───────────────────────────
  // Panels: CG_U, CG_L
  'MB-65': ['CG_U','CG_L'],         // Full 20x31
  'MB-66': ['CG_U'],                // Upper
  'MB-67': ['CG_L'],                // Lower

  // ── Group 16: South Bopal – Sobo Center ──────────────────────────────────
  // Panels: SOBO_U, SOBO_L
  'MB-68': ['SOBO_U','SOBO_L'],     // Full 30x31
  'MB-69': ['SOBO_U'],              // Upper
  'MB-70': ['SOBO_L'],              // Lower

  // ── Group 17: 200ft Ring Road Junction – Left ─────────────────────────────
  // Panels: RRL_U, RRL_L
  'MB-71': ['RRL_U','RRL_L'],       // Full 30x31
  'MB-72': ['RRL_U'],               // Upper
  'MB-73': ['RRL_L'],               // Lower

  // ── Group 18: 200ft Ring Road Junction – Right ────────────────────────────
  // Panels: RRR_U, RRR_L
  'MB-74': ['RRR_U','RRR_L'],       // Full 30x31
  'MB-75': ['RRR_U'],               // Upper
  'MB-76': ['RRR_L'],               // Lower

  // ── Group 19: Billionaire Street – Nr. Karnavati Club ────────────────────
  // Panels: BILL_U, BILL_L
  'MB-77': ['BILL_U','BILL_L'],     // Full 30x31
  'MB-78': ['BILL_U'],              // Upper
  'MB-79': ['BILL_L'],              // Lower

  // ── Group 20: Bopal Road – Nr. TRP Mall (vertical split) ─────────────────
  // Panels: TRP_U, TRP_L
  'MB-80': ['TRP_U','TRP_L'],       // Full 30x31
  'MB-81': ['TRP_U'],               // Upper (Aksardhara)
  'MB-82': ['TRP_L'],               // Lower (Aksardhara)

  // ── Group 21: Sanathal Circle Bridge ─────────────────────────────────────
  // Panels: SAT_U, SAT_L
  'MB-85': ['SAT_U','SAT_L'],       // Full 40x41
  'MB-86': ['SAT_U'],               // Upper
  'MB-87': ['SAT_L'],               // Lower
};

// ─── Build reverse map: panel → sites ─────────────────────────────────────
const PANEL_TO_SITES = {};
for (const [code, panels] of Object.entries(SITE_PANELS)) {
  for (const panel of panels) {
    if (!PANEL_TO_SITES[panel]) PANEL_TO_SITES[panel] = [];
    PANEL_TO_SITES[panel].push(code);
  }
}

/**
 * Returns all site codes that share any physical panel with the given siteCode.
 * Includes the site itself. Returns empty array for standalone sites.
 */
export function getOverlappingSiteCodes(siteCode) {
  const code = String(siteCode || '').toUpperCase().trim();
  const myPanels = SITE_PANELS[code];
  if (!myPanels || myPanels.length === 0) return [code];

  const affected = new Set();
  for (const panel of myPanels) {
    const siblings = PANEL_TO_SITES[panel] || [];
    for (const s of siblings) affected.add(s);
  }
  return Array.from(affected);
}

/**
 * Given a list of directly booked site codes (from active campaigns),
 * returns a Map of { siteCode → reason_string } for ALL occupied sites,
 * including both direct bookings and linked sites.
 *
 * @param {string[]} bookedCodes - site_codes that have active campaigns
 * @returns {Map<string, string>} - siteCode → reason (e.g. 'Direct' or 'Linked via MB-06')
 */
export function computeAllOccupied(bookedCodes) {
  const result = new Map();
  for (const code of bookedCodes) {
    const upper = String(code || '').toUpperCase().trim();
    if (!upper) continue;
    // Direct booking
    if (!result.has(upper)) result.set(upper, 'Direct');
    // Linked sites
    const linked = getOverlappingSiteCodes(upper);
    for (const linked_code of linked) {
      if (linked_code !== upper && !result.has(linked_code)) {
        result.set(linked_code, `Linked via ${upper}`);
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

/** Return structured conflict summary for UI or API warning */
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

/**
 * Automatically creates/updates linked campaign records in the campaigns table
 * for all physically overlapping sites.
 *
 * If a single site (e.g. MB-02) is booked, it creates linked campaign entries for
 * combined sites (MB-04, MB-05, MB-06).
 *
 * Vice versa: If a combined site (e.g. MB-06) is booked, it creates linked campaign
 * entries for single sites (MB-02, MB-03) and other overlapping combined sites (MB-04, MB-05).
 */
export async function syncLinkedCampaigns(q, primaryCampaignId) {
  try {
    const rows = await q('SELECT * FROM campaigns WHERE id=? LIMIT 1', [primaryCampaignId]);
    if (!rows[0]) return;
    const primary = rows[0];

    // Do not cascade if this is already an auto-booked linked campaign
    if (String(primary.parent_campaign || '').startsWith('LINKED:')) return;

    const rawCode = canonicalSiteCode(primary.site_code);
    if (!rawCode) return;

    const parentTag = `LINKED:${primary.id}`;

    // If primary is archived or soft-deleted, archive the linked campaigns
    if (primary.record_status === 'archived') {
      await q(`UPDATE campaigns SET record_status='archived', updated_at=NOW() WHERE parent_campaign=?`, [parentTag]);
      return;
    }

    // Get all overlapping sites (excluding self)
    const overlapping = getOverlappingSiteCodes(rawCode).filter(c => c !== rawCode);
    if (overlapping.length === 0) {
      await q(`DELETE FROM campaigns WHERE parent_campaign=?`, [parentTag]);
      return;
    }

    // Clean up any linked records that are no longer part of overlapping
    const placeholders = overlapping.map(() => '?').join(',');
    await q(`DELETE FROM campaigns WHERE parent_campaign=? AND site_code NOT IN (${placeholders})`, [parentTag, ...overlapping]);


    // Sync each linked site
    for (const linkedCode of overlapping) {
      const isComb = isCombinedSite(linkedCode);
      const displayTitle = primary.campaign_name || primary.display || primary.brand || 'Standard Display';
      const notesText = `Auto-booked: Linked to ${primary.site_code} booking #${primary.booking_code || primary.id}`;
      const bookingCode = `MB-BK-LNK-${primary.id}-${linkedCode}`;

      const existing = await q(
        `SELECT id FROM campaigns WHERE parent_campaign=? AND site_code=? LIMIT 1`,
        [parentTag, linkedCode]
      );

      if (existing[0]) {
        await q(`UPDATE campaigns SET
          client=?,
          brand=?,
          campaign_name=?,
          booking_date=?,
          start_date=?,
          end_date=?,
          revenue=0,
          vendor_cost=0,
          printing_cost=0,
          mounting_cost=0,
          electricity_cost=0,
          other_cost=0,
          notes=?,
          record_status='active',
          updated_at=NOW()
          WHERE id=?`, [
          primary.client || '',
          primary.brand || '',
          displayTitle,
          primary.booking_date || null,
          primary.start_date || null,
          primary.end_date || null,
          notesText,
          existing[0].id
        ]);
      } else {
        await q(`INSERT INTO campaigns (
          booking_code, parent_campaign, site_id, site_code, client, brand, campaign_name,
          booking_date, start_date, end_date,
          revenue, vendor_cost, printing_cost, mounting_cost, electricity_cost, other_cost,
          notes, record_status, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0, 0, 0, 0, 0, ?, 'active', NOW(), NOW())`, [
          bookingCode,
          parentTag,
          primary.site_id || null,
          linkedCode,
          primary.client || '',
          primary.brand || '',
          displayTitle,
          primary.booking_date || null,
          primary.start_date || null,
          primary.end_date || null,
          notesText
        ]);
      }
    }
  } catch (err) {
    console.error('[syncLinkedCampaigns] Error:', err.message);
  }
}

export async function deleteLinkedCampaigns(q, primaryCampaignId) {
  try {
    await q(`DELETE FROM campaigns WHERE parent_campaign=?`, [`LINKED:${primaryCampaignId}`]);
  } catch (err) {
    console.error('[deleteLinkedCampaigns] Error:', err.message);
  }
}

export async function batchDeleteLinkedCampaigns(q, primaryCampaignIds) {
  try {
    if (!primaryCampaignIds || primaryCampaignIds.length === 0) return;
    const tags = primaryCampaignIds.map(id => `LINKED:${id}`);
    const placeholders = tags.map(() => '?').join(',');
    await q(`DELETE FROM campaigns WHERE parent_campaign IN (${placeholders})`, tags);
  } catch (err) {
    console.error('[batchDeleteLinkedCampaigns] Error:', err.message);
  }
}

export async function syncAllLinkedCampaigns(q) {
  try {
    const primaries = await q(
      `SELECT id FROM campaigns WHERE record_status='active' AND (parent_campaign IS NULL OR parent_campaign NOT LIKE 'LINKED:%')`
    );
    for (const p of primaries) {
      await syncLinkedCampaigns(q, p.id);
    }
  } catch (err) {
    console.error('[syncAllLinkedCampaigns] Error:', err.message);
  }
}


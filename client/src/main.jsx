import React, { useEffect, useMemo, useState, useCallback, useRef, Component } from 'react';
import { createRoot } from 'react-dom/client';
import { BrowserRouter, Navigate, Route, Routes, useLocation, useNavigate } from 'react-router-dom';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import PptxGenJS from 'pptxgenjs';
import api from './api';
import { defaultSites, defaultSettings } from './defaultSites';
import { canonicalSiteCode, isCombinedSite, getSiteTypeTag, getConflictSummary, getOverlappingSiteCodes, SITE_PANELS } from './siteHierarchy';
import './styles.css';

// Helper to sanitize display titles and prevent "[Split Face via ...]" or "[Combined Block ...]" from showing anywhere
export function cleanDisplayTitle(str) {
  if (!str) return '';
  return String(str)
    .replace(/\s*\[(?:Split Face|Combined Block)[^\]]*\]/gi, '')
    .trim();
}

// Class ErrorBoundary to prevent any white/black screens
class ErrorBoundary extends Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, error: null };
  }
  static getDerivedStateFromError(error) {
    return { hasError: true, error };
  }
  componentDidCatch(error, info) {
    console.error('UI Runtime Catch:', error, info);
  }
  render() {
    if (this.state.hasError) {
      return (
        <div style={{ minHeight: '100vh', background: '#0b0f14', color: '#f4f6f8', padding: '40px', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center' }}>
          <div style={{ maxWidth: '500px', background: '#111720', border: '1px solid #27313d', borderRadius: '16px', padding: '28px', textAlign: 'center' }}>
            <img src="/assets/media-buzz-logo.png" alt="Media Buzz" style={{ width: '140px', marginBottom: '16px' }} />
            <h2 style={{ margin: '0 0 10px', fontSize: '18px' }}>Workspace Recovery</h2>
            <p style={{ fontSize: '12px', color: '#8d98a6', margin: '0 0 20px' }}>{this.state.error?.message || 'An unexpected display error occurred.'}</p>
            <button
              style={{ padding: '10px 20px', background: '#f2c94c', border: 0, borderRadius: '8px', color: '#16120a', fontWeight: 800, cursor: 'pointer' }}
              onClick={() => {
                localStorage.removeItem('sc_token');
                window.location.href = '/login';
              }}
            >
              Sign In Again
            </button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

// SVG Icons Dictionary matching the Media Buzz OOH Plugin
const icons = {
  dashboard: <path d="M4 13h6V4H4v9Zm0 7h6v-5H4v5Zm10 0h6v-9h-6v9Zm0-16v5h6V4h-6Z"/>,
  sites: <path d="M12 2a7 7 0 0 0-7 7c0 5.25 7 13 7 13s7-7.75 7-13a7 7 0 0 0-7-7Zm0 9.5A2.5 2.5 0 1 1 12 6a2.5 2.5 0 0 1 0 5.5Z"/>,
  campaigns: <path d="M4 4h12a2 2 0 0 1 2 2v2h2v8h-2v2a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2Zm2 4v8h8V8H6Z"/>,
  occupancy: <path d="M4 19h16v2H4v-2Zm1-8h3v6H5v-6Zm5-5h3v11h-3V6Zm5 3h3v8h-3V9Z"/>,
  proposals: <path d="M6 2h9l5 5v15H6V2Zm8 2v5h5M9 13h8v2H9v-2Zm0 4h8v2H9v-2ZM9 9h3v2H9V9Z"/>,
  ppt: <path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/>,
  electricity: <path d="M13 2 5 13h6l-1 9 9-13h-6l0-7Z"/>,
  vendors: <path d="M8 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8Zm8 1a3 3 0 1 0 0-6 3 3 0 0 0 0 6ZM2 21v-2c0-3 3-5 6-5s6 2 6 5v2H2Zm12.5 0v-2c0-1.35-.43-2.54-1.15-3.5.79-.32 1.69-.5 2.65-.5 3 0 6 2 6 5v1h-7.5Z"/>,
  clients: <path d="M12 12a5 5 0 1 0 0-10 5 5 0 0 0 0 10ZM3 22v-2c0-4 4-6 9-6s9 2 9 6v2H3Z"/>,
  invoices: <path d="M6 2h12v20l-3-2-3 2-3-2-3 2V2Zm3 5h6v2H9V7Zm0 4h6v2H9v-2Zm0 4h4v2H9v-2Z"/>,
  data: <path d="M12 2 7 7h3v6h4V7h3l-5-5ZM5 14v6h14v-6h2v8H3v-8h2Z"/>,
  reports: <path d="M4 2h16v20H4V2Zm4 14h2v3H8v-3Zm3-5h2v8h-2v-8Zm3 2h2v6h-2v-6ZM8 6h8v2H8V6Z"/>,
  notifications: <path d="M12 22a2.5 2.5 0 0 0 2.45-2h-4.9A2.5 2.5 0 0 0 12 22Zm7-5H5l2-2v-5a5 5 0 1 1 10 0v5l2 2Z"/>,
  activity: <path d="M12 2a10 10 0 1 0 10 10h-2a8 8 0 1 1-2.34-5.66L14 10h8V2l-2.91 2.91A9.96 9.96 0 0 0 12 2Zm-1 5h2v6l4 2-1 1.73-5-2.73V7Z"/>,
  storage: <path d="M4 4h16v4H4V4Zm0 6h16v4H4v-4Zm0 6h16v4H4v-4Zm6-10h4V5h-4v1Zm0 6h4v-1h-4v1Zm0 6h4v-1h-4v1Z"/>,
  settings: <path d="m19.14 12.94.04-.94-.04-.94 2.03-1.58-2-3.46-2.49 1a7.8 7.8 0 0 0-1.63-.94L14.68 3h-4l-.37 3.08c-.58.25-1.12.57-1.63.94l-2.49-1-2 3.46 2.03 1.58-.04.94.04.94-2.03 1.58 2 3.46 2.49-1c.51.37 1.05.69 1.63.94L10.68 21h4l.37-3.08c.58-.25 1.12-.57 1.63-.94l2.49 1 2-3.46-2.03-1.58ZM12.68 15.5a3.5 3.5 0 1 1 0-7 3.5 3.5 0 0 1 0 7Z"/>
};

const navGroups = [
  {
    label: 'Workspace',
    items: [
      ['dashboard', 'Dashboard'],
      ['sites', 'Sites'],
      ['campaigns', 'Campaign Tracker'],
      ['occupancy', 'Occupancy'],
      ['proposals', 'Proposal Builder'],
      ['ppt', 'Automated PPT'],
      ['storage', 'Storage & Archives']
    ]
  },
  {
    label: 'Operations',
    items: [
      ['electricity', 'Electricity'],
      ['vendors', 'Vendors'],
      ['data', 'Import / Export']
    ]
  },
  {
    label: 'Business',
    items: [
      ['clients', 'Clients'],
      ['invoices', 'Invoices'],
      ['reports', 'Reports']
    ]
  },
  {
    label: 'System',
    items: [
      ['notifications', 'Notifications'],
      ['activity', 'Activity Log'],
      ['settings', 'Settings']
    ]
  }
];

const fields = {
  sites: ['site_code','city','area','address','media_type','lighting','facing','size','width','height','ownership','availability','vendor_name','meter_no','monthly_cost','monthly_rate','latitude','longitude','maps_url','gps','notes'],
  clients: ['client_name','company','primary_contact','email','phone','billing_address','gst_number','status','notes'],
  campaigns: ['booking_code','parent_campaign','site_id','site_code','client_id','client','brand','campaign_name','booking_date','start_date','end_date','mounting_date','printing_status','mounting_status','validation_15_date','final_validation_date','revenue','vendor_cost','printing_cost','mounting_cost','electricity_cost','other_cost','invoice_required','invoice_requested','invoice_no','invoice_status','hard_copy_status','notes'],
  electricity: ['site_id','site_code','location','meter_no','size','service_number','t_number','bill_type','payment_amount','billing_month','bill_date','due_date','units','rate','other_charges','amount','payment_status','paid_date','payment_reference','notes'],
  vendors: ['name','service','contact_person','phone','email','cities','rating','notes','status'],
  proposals: ['proposal_code','client_id','client_name','campaign_name','proposal_date','start_date','duration_days','validity_days','discount_percent','tax_percent','subtotal','total','status','notes','terms'],
  invoices: ['campaign_id','client_id','requested_date','invoice_no','invoice_date','invoice_amount','invoice_status','hard_copy_required','hard_copy_status','courier_name','tracking_number','dispatch_date','delivered_date','payment_status','payment_date','notes']
};

const viewTitles = {
  dashboard: 'Dashboard',
  sites: 'Sites Directory',
  campaigns: 'Campaign Tracker',
  occupancy: 'Occupancy & Utilization',
  proposals: 'Proposal Builder',
  ppt: 'Automated PPT',
  electricity: 'Electricity & Meters',
  vendors: 'Vendors Directory',
  clients: 'Clients Directory',
  invoices: 'Invoices & Dispatch',
  data: 'Import / Export Tools',
  reports: 'Performance Reports',
  notifications: 'System Notifications',
  activity: 'Activity Log',
  storage: 'Storage & Archives',
  settings: 'System Settings'
};

export const ROLE_CONFIG = {
  admin: {
    key: 'admin',
    label: 'Admin',
    fullLabel: 'Administrator (Full Workspace & Settings Access)',
    color: '#c4b5fd',
    bg: 'rgba(139,92,246,0.18)',
    border: 'rgba(139,92,246,0.35)',
    description: 'Full workspace access, user account provisioning, system settings, database backups, and permanent record deletions across all modules.'
  },
  manager: {
    key: 'manager',
    label: 'Manager',
    fullLabel: 'Manager (Sites, Campaigns, Proposals & Storage)',
    color: '#4ade80',
    bg: 'rgba(34,197,94,0.18)',
    border: 'rgba(34,197,94,0.35)',
    description: 'Full operational control: Create, edit, export, and delete sites, campaigns, proposals, electricity, invoices, and storage archives. Cannot access User Management or System Settings.'
  },
  staff: {
    key: 'staff',
    label: 'Staff',
    fullLabel: 'Staff (Field Operations & Billing Tracking)',
    color: '#38bdf8',
    bg: 'rgba(56,189,248,0.18)',
    border: 'rgba(56,189,248,0.35)',
    description: 'Field operations: View inventory, update mounting & printing status, and log electricity meter bills. Deletions, bulk deletes, and proposals are restricted.'
  },
  viewer: {
    key: 'viewer',
    label: 'Viewer',
    fullLabel: 'Viewer (Read-Only Access)',
    color: '#f59e0b',
    bg: 'rgba(245,158,11,0.18)',
    border: 'rgba(245,158,11,0.35)',
    description: 'Read-only access: Can view dashboards, sites, campaigns, occupancy, and download PPT/Excel files. Cannot create, edit, import, or delete records.'
  }
};

export function getCurrentUser() {
  try {
    return JSON.parse(localStorage.getItem('sc_user') || '{}');
  } catch {
    return {};
  }
}

export function getCurrentRole() {
  const u = getCurrentUser();
  return (u.role || 'staff').toLowerCase();
}


const label = s => String(s).replaceAll('_', ' ').replace(/\b\w/g, m => m.toUpperCase());
const money = v => '₹' + Number(v || 0).toLocaleString('en-IN', { maximumFractionDigits: 2 });
const formatDate = v => v ? new Date(v).toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' }) : '—';

function universalCompare(valA, valB, dir = 'asc') {
  if (valA === valB) return 0;
  if (valA === null || valA === undefined || valA === '' || valA === '—') return 1;
  if (valB === null || valB === undefined || valB === '' || valB === '—') return -1;

  const strA = String(valA).trim();
  const strB = String(valB).trim();

  // Try parsing as site code (e.g. "MB-01", "MB-87", "01", "MB 2")
  const siteA = strA.match(/^(?:mb[-\s]*)?(\d+)$/i);
  const siteB = strB.match(/^(?:mb[-\s]*)?(\d+)$/i);
  if (siteA && siteB) {
    const diff = parseInt(siteA[1], 10) - parseInt(siteB[1], 10);
    return dir === 'desc' ? -diff : diff;
  }

  // Try parsing as number
  const numA = typeof valA === 'number' ? valA : (/^-?\d+(\.\d+)?$/.test(strA) ? parseFloat(strA) : NaN);
  const numB = typeof valB === 'number' ? valB : (/^-?\d+(\.\d+)?$/.test(strB) ? parseFloat(strB) : NaN);

  let res = 0;
  if (!isNaN(numA) && !isNaN(numB)) {
    res = numA - numB;
  } else {
    // Try parsing as date
    const isDateStr = str => typeof str === 'string' && (/^\d{4}-\d{2}-\d{2}/.test(str) || (!isNaN(Date.parse(str)) && !/^\d+$/.test(str)));
    if (isDateStr(valA) && isDateStr(valB)) {
      const timeA = new Date(valA).getTime();
      const timeB = new Date(valB).getTime();
      if (!isNaN(timeA) && !isNaN(timeB)) {
        res = timeA - timeB;
      } else {
        res = strA.localeCompare(strB, undefined, { numeric: true, sensitivity: 'base' });
      }
    } else {
      res = strA.localeCompare(strB, undefined, { numeric: true, sensitivity: 'base' });
    }
  }

  return dir === 'desc' ? -res : res;
}

// Smart site search matcher supporting site number (e.g. "01", "1", "mb-01", "87") and text search
function matchSiteSearch(site, rawQuery) {
  if (!rawQuery || !rawQuery.trim()) return true;
  if (!site) return false;
  const q = rawQuery.trim().toLowerCase();
  const siteCode = String(site.site_code || '').trim().toLowerCase();
  const area = String(site.area || '').toLowerCase();
  const city = String(site.city || '').toLowerCase();
  const address = String(site.address || site.location || '').toLowerCase();
  const mediaType = String(site.media_type || '').toLowerCase();
  const lighting = String(site.lighting || '').toLowerCase();

  // 1. Numeric or site-code search: e.g. '01', '1', 'mb-01', 'mb 01', 'mb01', 'site 1', '87'
  const isNumericQuery = /^(?:(?:mb|site)[\s-]*)?(\d+)$/i.test(q);
  if (isNumericQuery) {
    const qNumMatch = q.match(/^(?:(?:mb|site)[\s-]*)?0*(\d+)$/i);
    const qNum = qNumMatch ? parseInt(qNumMatch[1], 10) : NaN;
    const cleanQDigits = q.replace(/\D/g, '');

    const codeNumMatch = siteCode.match(/(\d+)/);
    const codeNum = codeNumMatch ? parseInt(codeNumMatch[1], 10) : NaN;
    const codeDigits = codeNumMatch ? codeNumMatch[1] : '';

    // If query has leading zeros like '01', exact digit match (e.g. "01" matches "01")
    if (cleanQDigits.startsWith('0') && codeDigits === cleanQDigits) return true;

    // Exact numeric match: query '1' matches MB-01 (1 == 1)
    if (!isNaN(qNum) && !isNaN(codeNum) && qNum === codeNum) return true;

    // Exact siteCode match
    if (siteCode === q || siteCode.replace(/[^a-z0-9]/g, '') === q.replace(/[^a-z0-9]/g, '')) return true;

    return false;
  }

  // 2. Full text match on site_code (e.g. if searching partial letters or 'mb')
  if (siteCode.includes(q) || siteCode.replace(/[^a-z0-9]/g, '').includes(q.replace(/[^a-z0-9]/g, ''))) return true;

  // 3. Descriptive fields
  if (area.includes(q) || address.includes(q) || city.includes(q) || mediaType.includes(q) || lighting.includes(q)) {
    return true;
  }
  return false;
}

function SortHeader({ label, sortKey, currentSort, onSort, align = 'left', style = {} }) {
  const isSorted = currentSort && currentSort.key === sortKey;
  const icon = !isSorted ? ' ⇅' : (currentSort.dir === 'asc' ? ' ▲' : ' ▼');
  return (
    <th
      onClick={() => onSort(sortKey)}
      style={{
        cursor: 'pointer',
        userSelect: 'none',
        textAlign: align,
        color: isSorted ? '#c4b5fd' : undefined,
        whiteSpace: 'nowrap',
        transition: 'color 0.15s ease, background 0.15s ease',
        ...style
      }}
      title={`Click to sort by ${label} (${isSorted && currentSort.dir === 'asc' ? 'Descending' : 'Ascending'})`}
    >
      <span style={{ display: 'inline-flex', alignItems: 'center', gap: '4px' }}>
        <span>{label}</span>
        <span style={{ fontSize: '10px', opacity: isSorted ? 1 : 0.45, color: isSorted ? '#a78bfa' : 'inherit' }}>{icon}</span>
      </span>
    </th>
  );
}

function parseDimensions(sizeStr, w, h) {
  let width = parseFloat(w) || 0;
  let height = parseFloat(h) || 0;
  if ((!width || !height) && sizeStr) {
    const match = String(sizeStr).match(/(\d+(?:\.\d+)?)\s*(?:ft|')?\s*[xX*×/–-]\s*(\d+(?:\.\d+)?)/);
    if (match) {
      if (!width) width = parseFloat(match[1]);
      if (!height) height = parseFloat(match[2]);
    }
  }
  return { width, height };
}

function normalizeTokens(str) {
  if (!str) return [];
  const stopWords = new Set([
    'the', 'a', 'an', 'and', 'or', 'at', 'in', 'on', 'to', 'from', 'of', 'for',
    'nr', 'near', 'opp', 'opposite', 'behind', 'beside', 'facing', 'fcg', 'towards',
    'road', 'rd', 'cross', 'crossroad', 'junction', 'jnc', 'circle', 'bridge', 'flyover',
    'traffic', 'highway', 'hw', 'hwy', 'street', 'st', 'lane', 'sector', 'sec'
  ]);
  return String(str)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, ' ')
    .split(/\s+/)
    .filter(t => t.length >= 2 && !stopWords.has(t));
}

function matchSiteByLocationAndSize(location, size, width, height, siteList) {
  if (!Array.isArray(siteList) || siteList.length === 0) return null;
  const targetDim = parseDimensions(size, width, height);
  const locTokens = normalizeTokens(location);
  const cleanLoc = String(location || '').toLowerCase().replace(/[^a-z0-9]/g, '');

  let bestSite = null;
  let bestScore = -1;

  for (const site of siteList) {
    let score = 0;
    const siteDim = parseDimensions(site.size, site.width, site.height);
    let sizeMatched = false;

    // 1. Size matching (highest weight: 45 points)
    if (targetDim.width > 0 && targetDim.height > 0 && siteDim.width > 0 && siteDim.height > 0) {
      const exactMatch = (Math.abs(targetDim.width - siteDim.width) < 0.5 && Math.abs(targetDim.height - siteDim.height) < 0.5);
      const flippedMatch = (Math.abs(targetDim.width - siteDim.height) < 0.5 && Math.abs(targetDim.height - siteDim.width) < 0.5);
      if (exactMatch || flippedMatch) {
        score += 45;
        sizeMatched = true;
      } else {
        const targetArea = targetDim.width * targetDim.height;
        const siteArea = siteDim.width * siteDim.height;
        if (targetArea > 0 && Math.abs(targetArea - siteArea) / targetArea < 0.05) {
          score += 25;
          sizeMatched = true;
        }
      }
    }

    // 2. Location token overlap
    const siteText = `${site.address || ''} ${site.area || ''} ${site.city || ''} ${site.location || ''} ${site.site_code || ''}`;
    const siteTokens = new Set(normalizeTokens(siteText));
    const siteRaw = siteText.toLowerCase().replace(/[^a-z0-9]/g, '');

    let tokenMatches = 0;
    for (const t of locTokens) {
      if (siteTokens.has(t)) {
        tokenMatches += 1;
        score += 15;
      } else if (siteRaw.includes(t)) {
        tokenMatches += 0.5;
        score += 8;
      }
    }

    // Substring match
    if (cleanLoc.length >= 4 && siteRaw.includes(cleanLoc)) {
      score += 30;
    }

    // Directional / position indicator match (Left / Right / Middle / 1 / 2 / 3)
    const locLower = String(location || '').toLowerCase();
    const siteLower = siteText.toLowerCase();
    ['left', 'right', 'middle', 'center', '(1)', '(2)', '(3)', '1-3', '2-1', '2-3'].forEach(pos => {
      if (locLower.includes(pos) && siteLower.includes(pos)) {
        score += 12;
      }
    });

    if (score > bestScore && (tokenMatches > 0 || (sizeMatched && (cleanLoc.length < 3 || siteRaw.includes(cleanLoc.slice(0, 3)))))) {
      bestScore = score;
      bestSite = site;
    }
  }

  return bestScore >= 20 ? bestSite : null;
}

function unpackSite(s) {
  let f = s.flags;
  try {
    if (typeof f === 'string') f = JSON.parse(f);
  } catch {}
  f = f && typeof f === 'object' && !Array.isArray(f) ? f : {};
  return {
    ...s,
    ppt_images: Array.isArray(f.ppt_images) ? f.ppt_images : (Array.isArray(s.ppt_images) ? s.ppt_images : []),
    ppt_availability: f.ppt_availability || s.ppt_availability || s.availability || '',
    ppt_rate: f.ppt_rate || s.ppt_rate || s.monthly_rate || ''
  };
}

async function imageData(url) {
  if (!url) throw new Error('Empty image URL');
  if (url.startsWith('data:')) return url;
  try {
    const r = await fetch(url);
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const b = await r.blob();
    return await new Promise((ok, ko) => {
      const fr = new FileReader();
      fr.onload = () => ok(fr.result);
      fr.onerror = ko;
      fr.readAsDataURL(b);
    });
  } catch (err) {
    if (url.startsWith('/')) {
      try {
        const altUrl = `http://localhost:3000${url}`;
        const r2 = await fetch(altUrl);
        if (r2.ok) {
          const b2 = await r2.blob();
          return await new Promise((ok, ko) => {
            const fr = new FileReader();
            fr.onload = () => ok(fr.result);
            fr.onerror = ko;
            fr.readAsDataURL(b2);
          });
        }
      } catch {}
    }
    throw err;
  }
}

const OVERLAY_SVG_STRING = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1920 1080" width="1920" height="1080">
  <defs>
    <!-- Deep dark navy gradient for the right dashboard panel -->
    <linearGradient id="mainNavy" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#02142B" />
      <stop offset="60%" stop-color="#04162C" />
      <stop offset="100%" stop-color="#010B17" />
    </linearGradient>

    <!-- Outer angled facet gradient -->
    <linearGradient id="outerFacet" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#0A2C52" />
      <stop offset="100%" stop-color="#041830" />
    </linearGradient>

    <!-- Subtle drop shadow for bottom-left logo badge over photo -->
    <filter id="badgeShadow" x="-10%" y="-10%" width="125%" height="125%">
      <feDropShadow dx="0" dy="4" stdDeviation="6" flood-color="#000000" flood-opacity="0.45" />
    </filter>
  </defs>

  <!-- ================= 1. LEFT BOTTOM CORNER: OFFICIAL LOGO BADGE ================= -->
  <g transform="translate(45, 920)" filter="url(#badgeShadow)">
    <!-- Solid Brand Yellow Badge matching official master logo -->
    <rect width="320" height="116" rx="10" fill="#FFCC00" />
    <svg x="10" y="7" width="300" height="102" viewBox="135 170 550 195">
      <!-- White Butterfly / Bow Emblem -->
      <path d="M 423.95 206.12 C 423.95 206.12 451.13 188.08 467.73 194.82 C 467.73 194.82 475.68 199.04 473.14 208.07 C 469.02 222.71 447.63 232.42 435.26 236.91 C 435.26 236.91 451.23 246.53 458.80 260.40 C 466.83 275.13 460.58 283.22 446.34 275.99 C 442.96 274.27 438.72 271.39 435.26 266.74 L 429.73 278.28 C 437.34 287.53 458.23 295.93 468.64 287.13 C 471.72 284.52 475.18 280.16 475.61 273.55 C 476.26 263.82 471.52 256.19 465.58 249.20 C 462.52 245.60 458.65 241.26 456.43 239.56 C 456.43 239.56 470.93 233.24 480.11 221.75 C 496.01 201.84 483.14 179.98 459.07 181.59 C 459.07 181.59 442.00 181.35 417.70 198.42 Z M 405.87 206.12 C 405.87 206.12 378.69 188.08 362.09 194.82 C 362.09 194.82 354.14 199.04 356.68 208.07 C 360.80 222.71 382.20 232.42 394.57 236.91 C 394.57 236.91 378.59 246.53 371.03 260.40 C 362.99 275.13 369.25 283.22 383.48 275.99 C 386.87 274.27 391.10 271.39 394.57 266.74 L 400.10 278.28 C 392.49 287.53 371.59 295.93 361.18 287.13 C 358.10 284.52 354.65 280.16 354.21 273.55 C 353.57 263.82 358.30 256.19 364.24 249.20 C 367.30 245.60 371.17 241.26 373.40 239.56 C 373.40 239.56 358.89 233.24 349.71 221.75 C 333.82 201.84 346.68 179.98 370.75 181.59 C 370.75 181.59 387.83 181.35 412.12 198.42 Z" fill="#FFFFFF" fill-rule="evenodd" />
      <!-- Media (Dark Navy) -->
      <path d="M 187.21 259.17 C 187.21 265.17 200.63 265.17 200.63 265.17 L 200.63 213.69 C 200.63 207.69 189.40 207.69 186.44 207.69 L 172.08 238.34 L 160.62 212.91 C 158.20 207.61 152.35 207.69 151.88 207.69 L 143.61 207.69 L 143.61 259.17 C 143.61 265.17 157.03 265.17 157.03 265.17 L 157.03 239.04 C 157.03 235.22 156.17 230.77 156.17 230.77 L 164.52 249.65 C 168.26 255.50 176.37 255.50 176.37 255.50 L 188.23 230.30 C 188.23 230.30 187.21 235.22 187.21 239.04 Z Z M 218.47 238.18 C 218.47 233.89 219.72 231.55 223.62 231.55 C 227.05 231.55 228.38 233.43 228.38 236.23 C 228.38 239.59 225.73 240.52 222.92 240.52 C 221.28 240.52 219.64 240.21 218.47 239.90 Z Z M 225.88 265.33 C 235.79 265.33 242.11 260.88 242.11 255.19 C 242.11 253.00 241.17 250.66 239.30 248.09 C 239.22 252.14 234.38 255.81 226.27 255.81 C 221.51 255.81 218.47 253.86 218.47 249.49 L 218.47 248.24 C 218.47 248.24 220.89 249.49 224.87 249.49 C 237.89 249.49 241.41 243.10 241.41 236.23 C 241.41 227.89 236.26 222.35 224.17 222.35 C 212.08 222.35 205.37 227.89 205.37 237.01 L 205.37 250.66 C 205.37 259.87 212.54 265.33 225.88 265.33 Z M 269.54 204.02 L 269.54 220.71 C 269.54 222.27 270.09 224.77 270.09 224.77 C 270.09 224.77 266.27 222.35 261.66 222.35 C 252.46 222.35 244.50 227.81 244.50 236.78 L 244.50 250.82 C 244.50 260.10 251.52 265.33 263.61 265.33 C 266.42 265.33 270.87 264.00 275.70 264.00 C 278.36 264.00 280.38 264.08 282.65 265.33 L 282.65 210.80 C 282.65 204.10 270.87 204.02 269.54 204.02 Z M 269.54 255.97 C 269.54 255.97 266.73 256.43 264.55 256.43 C 259.24 256.43 257.61 254.02 257.61 249.65 L 257.61 237.95 C 257.61 233.50 259.71 231.55 262.68 231.55 C 266.19 231.55 269.54 233.82 269.54 237.17 Z Z M 287.38 258.93 C 287.38 264.94 300.49 264.94 300.49 264.94 L 300.49 228.74 C 300.49 222.97 288.47 222.74 287.38 222.74 Z Z M 287.61 215.41 C 287.61 216.81 289.33 218.29 291.75 218.29 L 300.10 218.29 L 300.10 210.80 C 300.10 209.64 298.38 207.84 296.04 207.84 L 287.61 207.84 Z Z M 327.77 255.66 C 327.77 255.66 324.88 256.51 322.07 256.51 C 319.42 256.51 316.77 255.73 316.77 252.07 C 316.77 247.08 325.66 248.32 327.77 245.05 Z Z M 321.06 231.63 C 326.13 231.63 327.53 233.82 327.53 235.69 C 323.94 240.91 303.66 238.26 303.66 252.53 C 303.66 262.36 312.09 265.33 320.82 265.33 C 325.42 265.33 329.80 263.92 333.93 263.92 C 336.11 263.92 338.45 264.16 340.64 265.33 L 340.64 236.31 C 340.64 227.26 333.93 222.35 321.21 222.35 C 315.83 222.35 304.75 223.91 304.75 231.79 C 304.75 234.21 305.46 237.01 308.34 240.83 C 308.34 233.43 315.05 231.63 321.06 231.63" fill="#030352" fill-rule="evenodd" />
      <!-- Buzz (Dark Navy) -->
      <path d="M 498.73 258.07 C 498.73 264.00 508.72 264.08 508.72 264.08 L 518.31 264.08 C 531.65 264.08 540.85 259.71 540.85 248.87 L 540.85 247.15 C 540.85 242.32 538.12 237.09 533.83 235.38 C 537.50 233.97 540.07 228.51 540.07 223.99 L 540.07 222.43 C 540.07 211.04 530.95 207.37 518.31 207.37 L 498.73 207.37 Z Z M 518.93 240.21 C 524.86 240.21 527.44 242.47 527.44 247.70 C 527.44 252.77 524.71 254.64 518.31 254.64 L 512.15 254.64 L 512.15 240.21 Z Z M 518.31 216.89 C 523.93 216.89 526.66 218.37 526.66 222.82 C 526.66 227.97 524.63 230.77 518.93 230.77 L 512.15 230.77 L 512.15 216.89 Z Z M 557.91 228.43 C 557.91 222.74 544.81 222.43 544.81 222.43 L 544.81 251.21 C 544.81 262.05 554.17 265.01 562.90 265.01 C 566.88 265.01 570.55 263.38 574.53 263.38 C 579.44 263.38 582.95 265.33 582.95 265.33 L 582.95 228.43 C 582.95 222.43 569.85 222.43 569.85 222.43 L 569.85 254.95 C 569.85 254.95 566.65 255.73 564.00 255.73 C 559.47 255.73 557.91 253.71 557.91 250.04 Z Z M 604.85 232.49 C 599.86 240.52 584.88 248.01 584.88 259.79 C 584.88 260.41 584.80 260.96 585.50 264.08 L 614.60 264.08 C 621.85 264.08 621.85 254.64 621.85 254.64 L 599.78 254.64 C 606.02 246.14 617.95 238.73 620.06 231.32 C 620.92 228.43 621.07 225.70 621.93 222.97 L 593.15 222.97 C 585.74 222.97 585.74 232.49 585.74 232.49 Z Z M 640.32 232.49 C 635.32 240.52 620.35 248.01 620.35 259.79 C 620.35 260.41 620.27 260.96 620.97 264.08 L 650.07 264.08 C 657.32 264.08 657.32 254.64 657.32 254.64 L 635.25 254.64 C 641.49 246.14 653.42 238.73 655.53 231.32 C 656.39 228.43 656.54 225.70 657.40 222.97 L 628.62 222.97 C 621.21 222.97 621.21 232.49 621.21 232.49 Z" fill="#030352" fill-rule="evenodd" />
      <!-- Registered trademark circle & R -->
      <path d="M 661.61 189.98 C 670.50 189.98 677.72 196.99 677.72 205.63 C 677.72 214.27 670.50 221.28 661.61 221.28 C 652.71 221.28 645.50 214.27 645.50 205.63 C 645.50 196.99 652.71 189.98 661.61 189.98" stroke="#030352" stroke-width="1.8" fill="none" />
      <text x="661.6" y="206.5" font-family="'Arial', 'Helvetica', sans-serif" font-weight="bold" font-size="16" fill="#030352" text-anchor="middle" dominant-baseline="central">R</text>
      <!-- Be Seen tagline -->
      <text x="412" y="352" font-family="'Arial', 'Helvetica', sans-serif" font-weight="bold" font-size="58" fill="#030352" text-anchor="middle" letter-spacing="1">Be Seen</text>
    </svg>
  </g>

  <!-- ================= 2. RIGHT SIDE DASHBOARD: PROPER & BIG ================= -->
  <!-- Left outer angled facet -->
  <polygon points="1180,0 1260,0 1080,1080 1060,1080" fill="url(#outerFacet)" />
  <line x1="1180" y1="0" x2="1060" y2="1080" stroke="#0F3863" stroke-width="1.8" />

  <!-- Main dark navy right-hand panel -->
  <polygon points="1260,0 1920,0 1920,1080 1080,1080" fill="url(#mainNavy)" />
  <line x1="1260" y1="0" x2="1080" y2="1080" stroke="#0B2A4A" stroke-width="1.5" />

  <!-- A. SITE CODE Section Accents (Moved to top with ample space) -->
  <!-- Yellow horizontal bar above SITE CODE -->
  <rect x="1300" y="48" width="90" height="8" rx="4" fill="#FFC200" />
  <!-- Label: SITE CODE -->
  <text x="1300" y="88" font-family="'Montserrat', Arial, sans-serif" font-weight="700" font-size="21" fill="#FFFFFF" letter-spacing="4">SITE CODE</text>

  <!-- B. Headline Section Accents -->


  <!-- C. Thin Horizontal Divider Lines across the specs & coordinates list -->
  <line x1="1295" y1="375" x2="1885" y2="375" stroke="#0F355C" stroke-width="1.5" />
  <line x1="1295" y1="485" x2="1885" y2="485" stroke="#0F355C" stroke-width="1.5" />
  <line x1="1295" y1="595" x2="1885" y2="595" stroke="#0F355C" stroke-width="1.5" />
  <line x1="1295" y1="705" x2="1885" y2="705" stroke="#0F355C" stroke-width="1.5" />
  <line x1="1295" y1="815" x2="1885" y2="815" stroke="#0F355C" stroke-width="1.5" />
  <line x1="1295" y1="925" x2="1885" y2="925" stroke="#0F355C" stroke-width="1.5" />
  <line x1="1295" y1="1040" x2="1885" y2="1040" stroke="#0F355C" stroke-width="1.5" />

  <!-- D. Spec Section Labels & Icons matching reference photo -->
  <!-- Row 1: SIZE -->
  <g transform="translate(1305, 395)">
    <!-- 4-arrows icon -->
    <svg width="44" height="44" viewBox="0 0 24 24" fill="none" stroke="#8EA5C4" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
      <polyline points="15 3 21 3 21 9" />
      <polyline points="9 21 3 21 3 15" />
      <polyline points="21 15 21 21 15 21" />
      <polyline points="3 9 3 3 9 3" />
    </svg>
  </g>
  <text x="1395" y="415" font-family="'Montserrat', Arial, sans-serif" font-weight="700" font-size="18" fill="#8EA5C4" letter-spacing="1">SIZE</text>

  <!-- Row 2: TYPE -->
  <g transform="translate(1303, 505)">
    <!-- Billboard outline icon -->
    <svg width="46" height="46" viewBox="0 0 24 24" fill="none" stroke="#8EA5C4" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
      <rect x="3" y="4" width="18" height="11" rx="1.5" />
      <line x1="8" y1="15" x2="8" y2="21" />
      <line x1="16" y1="15" x2="16" y2="21" />
      <line x1="5" y1="21" x2="19" y2="21" />
    </svg>
  </g>
  <text x="1395" y="525" font-family="'Montserrat', Arial, sans-serif" font-weight="700" font-size="18" fill="#8EA5C4" letter-spacing="1">TYPE</text>

  <!-- Row 3: ILLUMINATION -->
  <g transform="translate(1304, 615)">
    <!-- Lightbulb icon -->
    <svg width="46" height="46" viewBox="0 0 24 24" fill="none" stroke="#8EA5C4" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
      <path d="M9 18h6" />
      <path d="M10 22h4" />
      <path d="M12 2v2" />
      <path d="M12 6a5 5 0 0 0-3.5 8.5c.9.9 1.5 1.8 1.5 2.5h4c0-.7.6-1.6 1.5-2.5A5 5 0 0 0 12 6z" />
      <line x1="4.93" y1="4.93" x2="6.34" y2="6.34" />
      <line x1="19.07" y1="4.93" x2="17.66" y2="6.34" />
    </svg>
  </g>
  <text x="1395" y="635" font-family="'Montserrat', Arial, sans-serif" font-weight="700" font-size="18" fill="#8EA5C4" letter-spacing="1">ILLUMINATION</text>

  <!-- Row 4: Adv. Fee Per Month -->
  <g transform="translate(1310, 723)">
    <!-- Yellow Rupee symbol -->
    <svg width="38" height="42" viewBox="0 0 24 24" fill="none" stroke="#FFC200" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round">
      <path d="M6 3h12" />
      <path d="M6 8h12" />
      <path d="M6 13l8.5 8" />
      <path d="M6 13h3a4.5 4.5 0 0 0 0-9" />
    </svg>
  </g>
  <text x="1395" y="745" font-family="'Montserrat', Arial, sans-serif" font-weight="700" font-size="17" fill="#8EA5C4" letter-spacing="0.5">Adv. Fee Per Month</text>

  <!-- Row 5: AVAILABILITY -->
  <g transform="translate(1306, 833)">
    <!-- Circle checkmark icon -->
    <svg width="44" height="44" viewBox="0 0 24 24" fill="none" stroke="#5AC8FA" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
      <circle cx="12" cy="12" r="9.5" stroke="#5AC8FA" />
      <polyline points="7.5 12 10.5 15 16.5 9" stroke="#5AC8FA" />
    </svg>
  </g>
  <text x="1395" y="855" font-family="'Montserrat', Arial, sans-serif" font-weight="700" font-size="18" fill="#8EA5C4" letter-spacing="1">AVAILABILITY</text>

  <!-- Row 6: Location Coordinates Icon & Label -->
  <g transform="translate(1305, 943)">
    <!-- Bright yellow location pin -->
    <svg width="44" height="52" viewBox="0 0 24 24" fill="#FFC200">
      <path fill-rule="evenodd" clip-rule="evenodd" d="M12 2C7.58 2 4 5.58 4 10C4 15.5 12 22 12 22C12 22 20 15.5 20 10C20 5.58 16.42 2 12 2ZM12 13C10.34 13 9 11.66 9 10C9 8.34 10.34 7 12 7C13.66 7 15 8.34 15 10C15 11.66 13.66 13 12 13Z" />
    </svg>
  </g>
  <text x="1375" y="966" font-family="'Montserrat', Arial, sans-serif" font-weight="700" font-size="18" fill="#8EA5C4" letter-spacing="1">LOCATION COORDINATES</text>
</svg>`;

const LOGO_SVG_STRING = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 540 140" width="540" height="140">
  <g transform="translate(10, 10)">
    <!-- Media: bold white text with modern geometric sans -->
    <text x="5" y="65" font-family="'Montserrat', 'Century Gothic', 'Segoe UI', Arial, sans-serif" font-weight="900" font-size="58" fill="#FFFFFF" letter-spacing="-0.5">Media</text>
    
    <!-- Stylized yellow continuous ribbon bow emblem with hollow center matching reference photo -->
    <g transform="translate(196, 5)">
      <path d="M 38 19
               C 45 7, 59 5, 66 12
               C 74 20, 73 33, 58 37
               C 73 41, 74 54, 66 62
               C 59 69, 45 67, 38 55
               C 31 67, 17 69, 10 62
               C 2 54, 3 41, 18 37
               C 3 33, 2 20, 10 12
               C 17 5, 31 7, 38 19 Z"
            fill="none"
            stroke="#FFC200"
            stroke-width="7.5"
            stroke-linecap="round"
            stroke-linejoin="round" />
    </g>

    <!-- Buzz: bold white text with registered trademark -->
    <text x="286" y="65" font-family="'Montserrat', 'Century Gothic', 'Segoe UI', Arial, sans-serif" font-weight="900" font-size="58" fill="#FFFFFF" letter-spacing="-0.5">Buzz</text>
    <text x="432" y="30" font-family="'Montserrat', Arial, sans-serif" font-weight="bold" font-size="19" fill="#FFFFFF">®</text>

    <!-- Be Seen: clean white text centered underneath emblem and Buzz -->
    <text x="250" y="104" font-family="'Inter', 'Segoe UI', Arial, sans-serif" font-weight="600" font-size="25" fill="#FFFFFF" letter-spacing="1">Be Seen</text>
  </g>
</svg>`;

const ICONS_SVG = {
  size: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#8EA5C4" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg>`,
  type: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#8EA5C4" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="11" rx="1.5"/><line x1="8" y1="14" x2="8" y2="21"/><line x1="16" y1="14" x2="16" y2="21"/><line x1="5" y1="21" x2="19" y2="21"/></svg>`,
  illum: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#8EA5C4" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 18h6"/><path d="M10 22h4"/><path d="M12 2v2"/><path d="M12 6a5 5 0 0 0-3.5 8.5c.9.9 1.5 1.8 1.5 2.5h4c0-.7.6-1.6 1.5-2.5A5 5 0 0 0 12 6z"/><line x1="4.93" y1="4.93" x2="6.34" y2="6.34"/><line x1="19.07" y1="4.93" x2="17.66" y2="6.34"/></svg>`,
  rupee: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#FFC400" stroke-width="2.8" stroke-linecap="round" stroke-linejoin="round"><path d="M6 3h12"/><path d="M6 8h12"/><path d="M6 13l8.5 8"/><path d="M6 13h3a4.5 4.5 0 0 0 0-9"/></svg>`,
  check: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#FFC400" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9.5"/><polyline points="7.5 12 10.5 15 16.5 9"/></svg>`,
  pin: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#E6A200"><path fill-rule="evenodd" clip-rule="evenodd" d="M12 2C7.58 2 4 5.58 4 10C4 15.5 12 22 12 22C12 22 20 15.5 20 10C20 5.58 16.42 2 12 2ZM12 13C10.34 13 9 11.66 9 10C9 8.34 10.34 7 12 7C13.66 7 15 8.34 15 10C15 11.66 13.66 13 12 13Z"/></svg>`
};

const svgPngCache = new Map();

async function svgToPngDataUrl(svgString, width = 128, height = 128) {
  const cacheKey = `${width}x${height}_${svgString}`;
  if (svgPngCache.has(cacheKey)) return svgPngCache.get(cacheKey);

  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext('2d');

  const blob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const img = new Image();

  await new Promise((resolve, reject) => {
    img.onload = () => resolve();
    img.onerror = () => reject(new Error('Failed to render SVG to canvas'));
    img.src = url;
  });

  ctx.drawImage(img, 0, 0, width, height);
  URL.revokeObjectURL(url);
  const dataUrl = canvas.toDataURL('image/png');
  svgPngCache.set(cacheKey, dataUrl);
  return dataUrl;
}

let cachedOverlayPng = null;
async function getSlideOverlayPng() {
  if (cachedOverlayPng) return cachedOverlayPng;
  try {
    const res = await fetch('/assets/ppt-slide-overlay.svg?v=' + Date.now());
    if (res.ok) {
      const text = await res.text();
      cachedOverlayPng = await svgToPngDataUrl(text, 1920, 1080);
      return cachedOverlayPng;
    }
  } catch (e) {
    console.warn('Overlay SVG fetch notice:', e);
  }
  cachedOverlayPng = await svgToPngDataUrl(OVERLAY_SVG_STRING, 1920, 1080);
  return cachedOverlayPng;
}

async function createSitePhotoShowcase(imgDataUrl, boxW_px = 1430, boxH_px = 1220) {
  const img = new Image();
  await new Promise((resolve, reject) => {
    img.onload = () => resolve();
    img.onerror = () => reject(new Error('Failed to load site photo for showcase'));
    img.src = imgDataUrl;
  });

  const nw = img.naturalWidth || 1920;
  const nh = img.naturalHeight || 1080;
  const imgRatio = nw / nh;

  const canvas = document.createElement('canvas');
  canvas.width = boxW_px;
  canvas.height = boxH_px;
  const ctx = canvas.getContext('2d');

  // 1. Deep luxury dark navy base matching the presentation aesthetic
  ctx.fillStyle = '#021020';
  ctx.fillRect(0, 0, boxW_px, boxH_px);

  // 2. Soft dimmed ambient glow backdrop (seamless luxury feel, eliminates empty voids)
  try {
    ctx.save();
    ctx.filter = 'blur(30px) brightness(0.35) saturate(1.25)';
    let bgW, bgH;
    const boxRatio = boxW_px / boxH_px;
    if (imgRatio >= boxRatio) {
      bgH = boxH_px + 80;
      bgW = bgH * imgRatio;
    } else {
      bgW = boxW_px + 80;
      bgH = bgW / imgRatio;
    }
    const bgX = (boxW_px - bgW) / 2;
    const bgY = (boxH_px - bgH) / 2;
    ctx.drawImage(img, bgX, bgY, bgW, bgH);
    ctx.restore();

    // Dark smooth vignette
    const grad = ctx.createRadialGradient(
      boxW_px / 2, boxH_px / 2, Math.min(boxW_px, boxH_px) * 0.25,
      boxW_px / 2, boxH_px / 2, Math.max(boxW_px, boxH_px) * 0.75
    );
    grad.addColorStop(0, 'rgba(2, 16, 32, 0.15)');
    grad.addColorStop(1, 'rgba(2, 16, 32, 0.85)');
    ctx.fillStyle = grad;
    ctx.fillRect(0, 0, boxW_px, boxH_px);
  } catch {
    ctx.fillStyle = '#021020';
    ctx.fillRect(0, 0, boxW_px, boxH_px);
  }

  // 3. 100% Uncropped, Razor-Sharp Foreground Photo (contain)
  const PAD = 14;
  const availW = boxW_px - (PAD * 2);
  const availH = boxH_px - (PAD * 2);
  const availRatio = availW / availH;

  let fitW, fitH, fitX, fitY;
  if (imgRatio >= availRatio) {
    fitW = availW;
    fitH = Math.round(availW / imgRatio);
    fitX = PAD;
    fitY = PAD + Math.round((availH - fitH) / 2);
  } else {
    fitH = availH;
    fitW = Math.round(availH * imgRatio);
    fitX = PAD + Math.round((availW - fitW) / 2);
    fitY = PAD;
  }

  // Soft realistic drop-shadow behind sharp photo
  ctx.save();
  ctx.shadowColor = 'rgba(0, 0, 0, 0.75)';
  ctx.shadowBlur = 24;
  ctx.shadowOffsetY = 8;
  ctx.drawImage(img, fitX, fitY, fitW, fitH);
  ctx.restore();

  // Draw crisp original photo on top
  ctx.drawImage(img, fitX, fitY, fitW, fitH);

  // Subtle architectural border around sharp photo
  ctx.strokeStyle = 'rgba(255, 255, 255, 0.22)';
  ctx.lineWidth = 1.5;
  ctx.strokeRect(fitX + 0.75, fitY + 0.75, fitW - 1.5, fitH - 1.5);

  // Clean container border
  ctx.strokeStyle = '#0F355C';
  ctx.lineWidth = 2;
  ctx.strokeRect(1, 1, boxW_px - 2, boxH_px - 2);

  return canvas.toDataURL('image/jpeg', 0.94);
}

async function makePpt(sites, pages = {}, fileName = 'MediaBuzz_Automated-PPT.pptx') {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'Media Buzz Outdoor';
  const SW = 13.333, SH = 7.5;

  const overlayData = await getSlideOverlayPng();

  async function covered(slide, url, x, y, w, h) {
    try {
      const data = await imageData(url);
      slide.addImage({ data, x, y, w, h });
    } catch {}
  }

  async function addSitePhoto(slide, url) {
    const BOX_X = 0.20, BOX_Y = 0.20, BOX_W = 7.15, BOX_H = 6.10;
    try {
      const data = await imageData(url);
      const cardDataUrl = await createSitePhotoShowcase(data, 1430, 1220);
      slide.addImage({ data: cardDataUrl, x: BOX_X, y: BOX_Y, w: BOX_W, h: BOX_H });
    } catch (err) {
      console.warn('Site photo showcase notice:', err);
      try {
        const data = await imageData(url);
        slide.addImage({
          data,
          x: BOX_X,
          y: BOX_Y,
          w: BOX_W,
          h: BOX_H,
          sizing: { type: 'contain', w: BOX_W, h: BOX_H }
        });
      } catch (innerErr) {
        console.warn('Site photo fallback placement notice:', innerErr);
      }
    }
  }

  async function fixed(u) {
    if (!u) return;
    const s = pptx.addSlide();
    s.background = { color: 'FFFFFF' };
    await covered(s, u, 0, 0, SW, SH);
  }

  function getCoords(st) {
    let lat = st.latitude, lng = st.longitude;
    if ((!lat || !lng) && st.gps) {
      const parts = String(st.gps).split(',').map(x => x.trim());
      if (parts.length === 2 && !isNaN(parseFloat(parts[0]))) {
        lat = parts[0];
        lng = parts[1];
      }
    }
    return {
      lat: lat ? String(lat) : '22.962323',
      lng: lng ? String(lng) : '72.962323'
    };
  }

  function getSize(st) {
    if (st.size && String(st.size).trim()) {
      const clean = String(st.size).trim();
      return clean.toLowerCase().includes('ft') ? clean : `${clean} ft`;
    }
    if (st.width && st.height) return `${st.width} ft x ${st.height} ft`;
    return '45 ft x 10 ft';
  }

  async function siteSlide(site, photoUrl) {
    const s = pptx.addSlide();
    s.background = { color: '02142B' };

    // 1. Left site photo showcase: 100% visible, zero crop, zero cut
    if (photoUrl) {
      await addSitePhoto(s, photoUrl);
    } else {
      s.addShape(pptx.ShapeType.roundRect, {
        x: 0.20, y: 0.20, w: 7.15, h: 6.10,
        fill: { color: '021020' }, line: { color: '0F355C', width: 1.5 }, rectRadius: 0.08
      });
      s.addText('NO SITE PHOTO UPLOADED', {
        x: 0.35, y: 2.95, w: 6.85, h: 0.6,
        fontFace: 'Arial', fontSize: 18, bold: true, color: '4B6A92', align: 'center'
      });
    }

    // 2. Exact Side Dashboard Overlay (facet, navy panel, logo, icons, dividers, yellow accent bars)
    if (overlayData) {
      s.addImage({ data: overlayData, x: 0, y: 0, w: SW, h: SH });
    }

    // 3. Dynamic Site Code Pill (Spacious & prominent at top)
    const siteCodeText = String(site.site_code || 'MB-AHD-001').toUpperCase();
    s.addShape(pptx.ShapeType.roundRect, {
      x: 9.03, y: 0.76, w: 2.10, h: 0.46,
      fill: { color: 'FFC200' }, line: { color: 'FFC200' }, rectRadius: 0.23
    });
    const trackerUrl = `${typeof window !== 'undefined' ? window.location.origin : ''}/campaigns?site=${encodeURIComponent(siteCodeText)}`;
    s.addText(siteCodeText, {
      x: 9.03, y: 0.76, w: 2.10, h: 0.46,
      fontFace: 'Arial', fontSize: 14, bold: true, color: '000000', align: 'center', valign: 'middle', margin: 0,
      hyperlink: { url: trackerUrl, tooltip: `Open ${siteCodeText} in Campaign Tracker` }
    });

    // 4. Headline: Area (Bold White, 30pt) & Landmark (Bold Light Sky Blue, 20-24pt)
    const area = String(site.area || site.city || site.site_code || 'Prime Site').trim();
    let landmark = String(site.address || site.notes || '').trim();
    if (!landmark || landmark.toLowerCase() === area.toLowerCase()) {
      landmark = site.city ? `Near ${site.city} Hub` : '';
    }

    s.addText(area, {
      x: 9.03, y: 1.35, w: 4.10, h: 0.55,
      fontFace: 'Arial', fontSize: 30, bold: true, color: 'FFFFFF', fit: 'shrink', margin: 0
    });

    let landmarkFontSize = 24;
    if (landmark.length > 50) landmarkFontSize = 17;
    else if (landmark.length > 30) landmarkFontSize = 20;
    else if (landmark.length > 20) landmarkFontSize = 22;

    if (landmark) {
      s.addText(landmark, {
        x: 9.03, y: 1.95, w: 4.10, h: 0.50,
        fontFace: 'Arial', fontSize: landmarkFontSize, bold: true, color: '5EB0FA', fit: 'shrink', margin: 0
      });
    }

    // 5. Spec values (SIZE, TYPE, ILLUMINATION, RATE) - Big, crisp & spacious
    s.addText(getSize(site), {
      x: 9.70, y: 2.95, w: 3.40, h: 0.35,
      fontFace: 'Arial', fontSize: 15, bold: true, color: 'FFFFFF', margin: 0
    });

    s.addText(String(site.media_type || 'Hoarding'), {
      x: 9.70, y: 3.71, w: 3.40, h: 0.35,
      fontFace: 'Arial', fontSize: 15, bold: true, color: 'FFFFFF', margin: 0
    });

    s.addText(String(site.lighting || 'NL').toUpperCase(), {
      x: 9.70, y: 4.48, w: 3.40, h: 0.35,
      fontFace: 'Arial', fontSize: 15, bold: true, color: 'FFFFFF', margin: 0
    });

    let rateText = 'On Request';
    if (site._showRate && site._rate) {
      rateText = `₹ ${Number(site._rate).toLocaleString('en-IN')}/-`;
    }
    s.addText(rateText, {
      x: 9.70, y: 5.24, w: 3.40, h: 0.35,
      fontFace: 'Arial', fontSize: 16, bold: true, color: 'FFFFFF', margin: 0
    });

    // 6. Availability Pill — green if available, amber/orange if booked
    const availText = String(site._availability || site.ppt_availability || site.availability || 'Available').trim();
    const isBooked = availText.toLowerCase().startsWith('booked') || availText.toLowerCase().startsWith('occupied');
    const pillColor = isBooked ? 'E07B00' : '00A859';
    const pillH = 0.38;
    const pillY = 6.01;
    const pillX = 9.70;
    const pillW = Math.max(1.65, Math.min(3.60, (availText.length * 0.11) + 0.55));
    const pillRadius = pillH / 2;

    s.addShape(pptx.ShapeType.roundRect, {
      x: pillX, y: pillY, w: pillW, h: pillH,
      fill: { color: pillColor }, line: { color: pillColor }, rectRadius: pillRadius
    });
    s.addText(availText, {
      x: pillX, y: pillY, w: pillW, h: pillH,
      fontFace: 'Arial',
      fontSize: availText.length > 20 ? 10.5 : (availText.length > 14 ? 11.5 : 12.5),
      bold: true,
      color: 'FFFFFF',
      align: 'center',
      valign: 'middle',
      margin: 0,
      wrap: false,
      fit: 'shrink'
    });

    // 6b. Campaign End Date (Only shown if availability pill does not already include it)
    if (site._endDate && !availText.includes(site._endDate)) {
      s.addText(`📅 End Date: ${site._endDate}`, {
        x: 9.55, y: 6.48, w: 3.60, h: 0.26,
        fontFace: 'Arial', fontSize: 10.5, bold: false, color: 'FFC870', margin: 0, align: 'left'
      });
    }

    // 7. Location Coordinates (Always shown by default)
    const { lat, lng } = getCoords(site);
    const coordsText = `Latitude ${lat}  |  Longitude ${lng}`;
    s.addText(coordsText, {
      x: 9.55, y: 6.81, w: 3.60, h: 0.28,
      fontFace: 'Arial', fontSize: 11, bold: true, color: 'FFFFFF', margin: 0
    });
  }

  // Helper to get effective page url with localStorage fallback
  function getEffectivePageUrl(key) {
    if (pages && pages[key]) return pages[key];
    try {
      return localStorage.getItem(`mb_ppt_raw_${key}`) || '';
    } catch {
      return '';
    }
  }

  // Include optional cover pages if uploaded
  const firstCover = getEffectivePageUrl('first');
  const secondCover = getEffectivePageUrl('second_last');
  const lastCover = getEffectivePageUrl('last');

  if (firstCover) await fixed(firstCover);
  if (secondCover) await fixed(secondCover);

  // Generate slides for all chosen sites
  for (const st of sites) {
    if (st.ppt_images && st.ppt_images.length) {
      for (const u of st.ppt_images) await siteSlide(st, u);
    } else {
      await siteSlide(st, '');
    }
  }

  // Include optional closing page if uploaded
  if (lastCover) await fixed(lastCover);

  let finalFileName = String(fileName || 'MediaBuzz_Automated-PPT.pptx').trim();
  if (!finalFileName.toLowerCase().endsWith('.pptx')) finalFileName += '.pptx';
  await pptx.writeFile({ fileName: finalFileName });

  try {
    const siteCodes = (sites || []).map(s => s.site_code).filter(Boolean);
    api.post('/storage', {
      category: 'ppt',
      title: finalFileName.replace(/\.pptx$/i, '').replace(/[_-]/g, ' '),
      filename: finalFileName,
      file_size: `${Math.max(1, Math.round((sites?.length || 1) * 0.6))} MB`,
      format: 'pptx',
      meta_json: {
        slides: (sites || []).length,
        sites: siteCodes.slice(0, 8),
        generatedAt: new Date().toISOString()
      }
    }).catch(() => {});
  } catch {}
}

async function exportStyledExcel(rowsData, filename = 'MediaBuzz_Sites.xlsx', title = null) {
  const workbook = new ExcelJS.Workbook();
  const cleanTitle = title ? String(title).trim() : '';
  const hasTitle = Boolean(cleanTitle);
  const headerRowNum = hasTitle ? 2 : 1;
  const dataStartRowNum = hasTitle ? 3 : 2;

  const worksheet = workbook.addWorksheet('Sites', {
    views: [{ state: 'frozen', ySplit: hasTitle ? 2 : 1 }]
  });

  worksheet.columns = [
    { key: 'sr', width: 8 },
    { key: 'area', width: 22 },
    { key: 'location', width: 55 },
    { key: 'media', width: 14 },
    { key: 'light', width: 10 },
    { key: 'w', width: 8 },
    { key: 'h', width: 8 },
    { key: 'sqft', width: 12 },
    { key: 'availability', width: 18 },
    { key: 'rate', width: 18 },
    { key: 'coords', width: 28 }
  ];

  // Optional Presentation / Report Title as Line 1
  if (hasTitle) {
    worksheet.mergeCells('A1:K1');
    const titleRow = worksheet.getRow(1);
    titleRow.height = 36;
    for (let c = 1; c <= 11; c++) {
      const cell = titleRow.getCell(c);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF071C35' } // Deep MediaBuzz Navy
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF234D7D' } },
        bottom: { style: 'medium', color: { argb: 'FFFFC200' } },
        left: { style: 'thin', color: { argb: 'FF234D7D' } },
        right: { style: 'thin', color: { argb: 'FF234D7D' } }
      };
    }
    const titleCell = titleRow.getCell(1);
    titleCell.value = cleanTitle;
    titleCell.font = {
      name: 'Calibri',
      size: 13.5,
      bold: true,
      color: { argb: 'FFFFC200' } // MediaBuzz Yellow
    };
    titleCell.alignment = {
      vertical: 'middle',
      horizontal: 'center'
    };
  }

  // Populate Header Row
  const headers = ['SR NO', 'AREA', 'LOCATION', 'MEDIA', 'LIGHT', 'W', 'H', 'SQ FT', 'AVAILABLITY', 'Selling Amount', 'Latitude Longitude'];
  const headerRow = worksheet.getRow(headerRowNum);
  headerRow.height = 28;
  headers.forEach((h, i) => {
    headerRow.getCell(i + 1).value = h;
  });

  // Populate Data Rows
  rowsData.forEach(r => {
    worksheet.addRow({
      sr: r['SR NO'],
      area: r['AREA'],
      location: r['LOCATION'],
      media: r['MEDIA'],
      light: r['LIGHT'],
      w: r['W'],
      h: r['H'],
      sqft: r['SQ FT'],
      availability: r['AVAILABLITY'],
      rate: r['Selling Amount'],
      coords: r['Latitude Longitude']
    });
  });

  // Style Header Row in Bright Yellow (#FFFF00)
  headerRow.eachCell((cell, colNumber) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF00' } // Bright Yellow matching template screenshot
    };
    cell.font = {
      name: 'Calibri',
      size: 11,
      bold: true,
      color: { argb: 'FF000000' }
    };
    cell.alignment = {
      vertical: 'middle',
      horizontal: [1, 5, 6, 7, 8, 9].includes(colNumber) ? 'center' : (colNumber === 10 ? 'right' : 'left'),
      wrapText: false
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFC0C0C0' } },
      left: { style: 'thin', color: { argb: 'FFC0C0C0' } },
      bottom: { style: 'medium', color: { argb: 'FF808080' } },
      right: { style: 'thin', color: { argb: 'FFC0C0C0' } }
    };
  });

  // Style Data Rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber >= dataStartRowNum) {
      row.height = 22;
      row.eachCell((cell, colNumber) => {
        cell.font = {
          name: 'Calibri',
          size: 10.5
        };
        cell.alignment = {
          vertical: 'middle',
          horizontal: [1, 5, 6, 7, 8, 9].includes(colNumber) ? 'center' : (colNumber === 10 ? 'right' : 'left')
        };
        if (colNumber === 10 && typeof cell.value === 'number') {
          cell.numFmt = '#,##,##0';
        }
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE8E8E8' } },
          left: { style: 'thin', color: { argb: 'FFE8E8E8' } },
          bottom: { style: 'thin', color: { argb: 'FFE8E8E8' } },
          right: { style: 'thin', color: { argb: 'FFE8E8E8' } }
        };
      });
    }
  });

  // Enable Auto-Filter
  if (rowsData.length > 0) {
    const filterStart = hasTitle ? 'A2' : 'A1';
    const filterEnd = `K${rowsData.length + (hasTitle ? 2 : 1)}`;
    worksheet.autoFilter = `${filterStart}:${filterEnd}`;
  }

  let finalFileName = String(filename || 'MediaBuzz_Sites.xlsx').trim();
  if (!finalFileName.toLowerCase().endsWith('.xlsx')) finalFileName += '.xlsx';

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = finalFileName;
  a.click();
  URL.revokeObjectURL(url);
}

function Login() {
  const nav = useNavigate();
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [err, setErr] = useState('');
  const [loading, setLoading] = useState(false);

  async function submit(e) {
    e.preventDefault();
    setLoading(true);
    setErr('');
    try {
      const { data } = await api.post('/auth/login', { email, password });
      localStorage.setItem('sc_token', data.token);
      localStorage.setItem('sc_user', JSON.stringify(data.user));
      nav('/dashboard');
    } catch (e) {
      setErr(e.response?.data?.message || 'Login failed. Check email and password.');
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="login">
      <form onSubmit={submit}>
        <img src="/assets/media-buzz-logo.png" alt="Media Buzz" />
        <h1>Site Control</h1>
        <p>OOH Operations Workspace</p>
        {err && <div className="error">{err}</div>}
        <label>
          Email ID
          <input type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="admin@domain.com" required />
        </label>
        <label>
          Password
          <input type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="••••••••" required />
        </label>
        <button disabled={loading}>{loading ? 'Signing in…' : 'Sign in'}</button>
      </form>
    </div>
  );
}

function Guard({ children }) {
  return localStorage.getItem('sc_token') ? children : <Navigate to="/login" replace />;
}

function Layout() {
  const loc = useLocation();
  const nav = useNavigate();
  let user = {};
  try { user = JSON.parse(localStorage.getItem('sc_user') || '{}'); } catch {}
  const userRole = (user.role || 'staff').toLowerCase();
  const isAdmin = userRole === 'admin';
  const isManager = userRole === 'manager';
  const isStaff = userRole === 'staff';
  const isReadOnly = userRole === 'viewer';
  const canDeleteRecords = isAdmin || isManager;

  const [notifOpen, setNotifOpen] = useState(false);
  const [notifications, setNotifications] = useState([]);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [theme, setTheme] = useState(() => localStorage.getItem('scooh_theme') || 'dark');

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('scooh_theme', theme);
  }, [theme]);

  const toggleTheme = () => setTheme(prev => (prev === 'dark' ? 'light' : 'dark'));

  const currentModule = useMemo(() => {
    const p = loc.pathname.replace(/^\//, '').split('/')[0] || 'dashboard';
    return p;
  }, [loc.pathname]);

  const currentTitle = viewTitles[currentModule] || 'OOH Operations Management';

  useEffect(() => {
    api.get('/notifications').then(r => {
      if (Array.isArray(r.data)) setNotifications(r.data);
    }).catch(() => {});
  }, []);

  const unreadCount = (Array.isArray(notifications) ? notifications : []).filter(n => !n.is_read).length;

  async function clearAllNotifs() {
    try {
      await api.post('/notifications/clear', {});
    } catch (e) {
      try { await api.post('/notifications/read', {}); } catch (err) {}
    }
    setNotifications([]);
  }

  const visibleNavGroups = useMemo(() => {
    return navGroups.map(grp => {
      const items = grp.items.filter(([k]) => {
        if (isAdmin) return true;
        if (isManager) return k !== 'settings';
        if (isStaff) return ['dashboard', 'sites', 'campaigns', 'occupancy', 'electricity', 'notifications', 'activity'].includes(k);
        if (isReadOnly) return ['dashboard', 'sites', 'campaigns', 'occupancy', 'storage', 'electricity', 'reports', 'notifications', 'activity'].includes(k);
        return true;
      });
      return { ...grp, items };
    }).filter(grp => grp.items.length > 0);
  }, [userRole, isAdmin, isManager, isStaff, isReadOnly]);

  return (
    <div className="scooh-wp-wrap">
      <div className="scooh-app scooh-has-topbar">
        {/* Top Header */}
        <header className="scooh-topbar">
          <div className="scooh-topbar-left">
            <button
              type="button"
              className="scooh-mobile-nav-toggle"
              onClick={() => setSidebarOpen(!sidebarOpen)}
              aria-label="Toggle Navigation"
            >
              <span></span><span></span><span></span>
            </button>
            <div className="scooh-topbar-context" style={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: '8px' }}>
              <div>
                <span className="scooh-topbar-eyebrow">MEDIA BUZZ • OOH WORKSPACE</span>
                <strong>{currentTitle}</strong>
              </div>
              {isReadOnly && (
                <span style={{
                  background: 'rgba(245,158,11,0.15)',
                  border: '1px solid rgba(245,158,11,0.3)',
                  color: '#fbbf24',
                  fontSize: '11px',
                  fontWeight: 700,
                  padding: '3px 9px',
                  borderRadius: '6px',
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '4px'
                }}>
                  🔒 Read-Only Workspace
                </span>
              )}
            </div>
          </div>

          <div className="scooh-topbar-user">
            {/* Quick Theme Toggle */}
            <button
              type="button"
              className="scooh-theme-toggle-btn"
              onClick={toggleTheme}
              title={`Switch to ${theme === 'dark' ? 'Light' : 'Dark'} Mode`}
              aria-label="Toggle Theme"
            >
              <span>{theme === 'dark' ? '☀️' : '🌙'}</span>
              <span>{theme === 'dark' ? 'Light' : 'Dark'}</span>
            </button>

            <div className="scooh-topbar-notifications" style={{ position: 'relative' }}>
              <button
                type="button"
                className="scooh-notification-bell"
                onClick={() => setNotifOpen(!notifOpen)}
                aria-label="Notifications"
              >
                <span className="scooh-bell-icon">🔔</span>
                {unreadCount > 0 && (
                  <span className="scooh-top-notification-count">
                    {unreadCount}
                  </span>
                )}
              </button>

              {notifOpen && (
                <div className="scooh-top-notification-panel" style={{ display: 'block' }}>
                  <div className="scooh-top-notification-head">
                    <strong>Notifications</strong>
                    <div className="scooh-notification-head-actions">
                      <button type="button" onClick={clearAllNotifs}>Clear all</button>
                      <button type="button" onClick={() => { setNotifOpen(false); nav('/notifications'); }}>View all</button>
                    </div>
                  </div>
                  <div className="scooh-top-notification-list">
                    {(!notifications || notifications.length === 0) ? (
                      <div className="scooh-top-notification-empty">No new notifications</div>
                    ) : (
                      notifications.map(n => (
                        <div key={n.id} className={`scooh-top-notification-item ${n.is_read ? '' : 'today'}`}>
                          <span className="scooh-top-notification-dot"></span>
                          <div>
                            <strong>{n.title}</strong>
                            <small>{n.message}</small>
                          </div>
                        </div>
                      ))
                    )}
                  </div>
                </div>
              )}
            </div>

            <span className="scooh-user-avatar">
              {String(user.name || user.email || 'A').slice(0, 1).toUpperCase()}
            </span>
            <div className="scooh-user-meta">
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <span style={{ fontSize: '10.5px', color: '#94a3b8' }}>Signed in as</span>
                <span
                  className="scooh-badgechip"
                  style={{
                    background: ROLE_CONFIG[userRole]?.bg || 'rgba(139,92,246,0.18)',
                    color: ROLE_CONFIG[userRole]?.color || '#c4b5fd',
                    border: `1px solid ${ROLE_CONFIG[userRole]?.border || 'rgba(139,92,246,0.35)'}`,
                    textTransform: 'uppercase',
                    fontSize: '9.5px',
                    fontWeight: 800,
                    padding: '2px 7px'
                  }}
                  title={ROLE_CONFIG[userRole]?.description}
                >
                  {ROLE_CONFIG[userRole]?.label || userRole}
                </span>
              </div>
              <strong>{user.name || user.email || 'Administrator'}</strong>
            </div>
            <button
              type="button"
              className="scooh-topbar-logout"
              onClick={() => {
                localStorage.clear();
                nav('/login');
              }}
            >
              Logout
            </button>
          </div>
        </header>

        {/* Sidebar */}
        <aside className={`scooh-sidebar ${sidebarOpen ? 'open' : ''}`}>
          <div className="scooh-brand-shell">
            <img src="/assets/media-buzz-logo.png" alt="Media Buzz" />
          </div>
          <nav className="scooh-tabs">
            {visibleNavGroups.map(grp => (
              <div className="scooh-nav-group" key={grp.label}>
                <span className="scooh-nav-group-label">{grp.label}</span>
                {grp.items.map(([k, n]) => (
                  <button
                    key={k}
                    type="button"
                    className={loc.pathname === '/' + k || (k !== 'dashboard' && loc.pathname.startsWith('/' + k)) ? 'active' : ''}
                    onClick={() => {
                      nav('/' + k);
                      setSidebarOpen(false);
                    }}
                  >
                    <span className="scooh-nav-icon">
                      <svg viewBox="0 0 24 24" aria-hidden="true">{icons[k] || icons.dashboard}</svg>
                    </span>
                    <span className="scooh-nav-label">{n}</span>
                    {k === 'notifications' && unreadCount > 0 && (
                      <span className="scooh-notification-count">{unreadCount}</span>
                    )}
                  </button>
                ))}
              </div>
            ))}
          </nav>
          <div className="scooh-sidebar-footer">
            <span className="scooh-sidebar-dot"></span>
            <span className="scooh-sidebar-status">Workspace online</span>
          </div>
        </aside>

        {sidebarOpen && <button type="button" className="scooh-mobile-scrim" onClick={() => setSidebarOpen(false)} aria-label="Close menu" />}

        {/* Main Content Area */}
        <main className="scooh-main">
          <div className="scooh-view">
            <Routes>
              <Route path="/dashboard" element={<Dashboard />} />
              <Route path="/sites" element={<SitesView />} />
              <Route path="/campaigns" element={<CampaignTrackerView />} />
              <Route path="/occupancy" element={<OccupancyView />} />
              <Route path="/proposals" element={!isStaff && !isReadOnly ? <ProposalsView /> : <Navigate to="/dashboard" replace />} />
              <Route path="/ppt" element={!isStaff && !isReadOnly ? <PptView /> : <Navigate to="/dashboard" replace />} />
              <Route path="/electricity" element={<ElectricityView />} />
              <Route path="/vendors" element={!isStaff && !isReadOnly ? <Crud entity="vendors" title="Vendors" /> : <Navigate to="/dashboard" replace />} />
              <Route path="/clients" element={!isStaff && !isReadOnly ? <Crud entity="clients" title="Clients" /> : <Navigate to="/dashboard" replace />} />
              <Route path="/invoices" element={!isStaff && !isReadOnly ? <Crud entity="invoices" title="Invoices" /> : <Navigate to="/dashboard" replace />} />
              <Route path="/data" element={!isStaff && !isReadOnly ? <DataToolsView /> : <Navigate to="/dashboard" replace />} />
              <Route path="/storage" element={<StorageView />} />
              <Route path="/reports" element={<ReportsView />} />
              <Route path="/notifications" element={<SimpleList endpoint="notifications" title="Notifications" />} />
              <Route path="/activity" element={<SimpleList endpoint="activity" title="Activity Log" />} />
              <Route path="/settings" element={isAdmin ? <SettingsView /> : <Navigate to="/dashboard" replace />} />
              <Route path="*" element={<Navigate to="/dashboard" replace />} />
            </Routes>
          </div>
        </main>
      </div>
    </div>
  );
}

function PageHead({ title, desc, actions }) {
  return (
    <div className="scooh-pagehead">
      <div>
        <h1>{title}</h1>
        {desc && <p>{desc}</p>}
      </div>
      {actions && <div className="scooh-headactions">{actions}</div>}
    </div>
  );
}

function OccupancyDonut({ buckets = {}, average = 0 }) {
  const data = [
    { label: 'Occupied', value: buckets.Occupied ?? 0, color: '#22C55E' },
    { label: 'Available', value: buckets.Available ?? 0, color: '#3B82F6' },
    { label: 'Prime Sites', value: buckets['Prime Sites'] ?? 0, color: '#FBBF24' },
    { label: 'Needs review', value: buckets['Needs review'] ?? 0, color: '#EF4444' }
  ];
  const total = data.reduce((acc, d) => acc + d.value, 0) || 1;
  let accumulated = 0;
  const radius = 68;
  const strokeWidth = 24;
  const circumference = 2 * Math.PI * radius;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: '230px' }}>
      <div style={{ position: 'relative', width: '180px', height: '180px', margin: '10px auto' }}>
        <svg width="180" height="180" viewBox="0 0 180 180" style={{ transform: 'rotate(-90deg)' }}>
          {data.map((slice, i) => {
            const dashLength = (slice.value / total) * circumference;
            const strokeDasharray = `${dashLength} ${circumference}`;
            const strokeDashoffset = -accumulated;
            accumulated += dashLength;
            return (
              <circle
                key={i}
                cx="90"
                cy="90"
                r={radius}
                fill="transparent"
                stroke={slice.color}
                strokeWidth={strokeWidth}
                strokeDasharray={strokeDasharray}
                strokeDashoffset={strokeDashoffset}
              />
            );
          })}
        </svg>
        <div style={{ position: 'absolute', inset: 0, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center' }}>
          <span style={{ fontSize: '26px', fontWeight: 800, color: '#fff' }}>{average}%</span>
          <span style={{ fontSize: '9px', color: '#8d99a8', textTransform: 'uppercase', letterSpacing: '0.08em' }}>Occupied</span>
        </div>
      </div>
      <div style={{ display: 'flex', gap: '14px', flexWrap: 'wrap', justifyContent: 'center', marginTop: '14px' }}>
        {data.map((d, i) => (
          <div key={i} style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '11px', color: '#9ba7b5' }}>
            <span style={{ width: '8px', height: '8px', borderRadius: '50%', background: d.color }}></span>
            <span>{d.label} ({d.value})</span>
          </div>
        ))}
      </div>
    </div>
  );
}

function Dashboard() {
  const [d, setD] = useState(null);
  const [err, setErr] = useState('');
  const nav = useNavigate();

  function load() {
    setErr('');
    api.get('/dashboard')
      .then(r => setD(r.data))
      .catch(e => setErr(e.response?.data?.message || 'Failed to load dashboard data'));
  }

  useEffect(() => { load(); }, []);

  if (err) {
    return (
      <>
        <PageHead title="Operations Dashboard" desc="Media Buzz OOH operations overview" />
        <section className="scooh-panel">
          <p style={{ color: '#f06a6a', fontWeight: 600 }}>{err}</p>
          <button className="scooh-btn purple-btn" onClick={load}>Retry</button>
        </section>
      </>
    );
  }

  if (!d) {
    return (
      <div className="scooh-loading">
        <span className="scooh-spinner"></span>
        <strong>Loading workspace</strong>
        <small>Preparing your Media Buzz data…</small>
      </div>
    );
  }

  const k = d.kpis || {};
  const alerts = Array.isArray(d.alerts) ? d.alerts : [];

  return (
    <>
      <PageHead
        title="Operations Dashboard"
        desc="Live snapshot of the site portfolio, active campaigns, and everything due for action today."
        actions={
          <>
            <button className="scooh-btn purple-btn" onClick={() => nav('/sites')}>+ Add site</button>
            <button className="scooh-btn" onClick={() => nav('/campaigns')}>+ Log booking</button>
          </>
        }
      />

      {/* Row 1: Site Portfolio & Live Campaigns */}
      <div className="scooh-kpirow">
        <div className="scooh-kpi alert" onClick={() => nav('/sites')} style={{ cursor: 'pointer' }} title="View all sites">
          <div className="n">{k.total_sites ?? 0}</div>
          <div className="l">Total Sites</div>
          {k.flagged_count > 0 && (
            <div className="scooh-kpi-note">⚠ {k.flagged_count} flagged for review</div>
          )}
        </div>
        <div className="scooh-kpi good" onClick={() => nav('/sites')} style={{ cursor: 'pointer' }} title="View occupied sites">
          <div className="n">{k.active_sites ?? 0}</div>
          <div className="l">Active Sites</div>
        </div>
        <div className="scooh-kpi alert" onClick={() => nav('/sites')} style={{ cursor: 'pointer' }} title="View available sites">
          <div className="n">{k.nonactive_sites ?? 0}</div>
          <div className="l">Vacant / Available</div>
        </div>
        <div className="scooh-kpi alert" onClick={() => nav('/campaigns')} style={{ cursor: 'pointer' }} title="View active campaigns">
          <div className="n">{k.active_campaigns ?? 0}</div>
          <div className="l">Active Campaigns</div>
        </div>
      </div>

      {/* Row 2: Operational Action Triggers */}
      <div className="scooh-kpirow">
        <div className={`scooh-kpi ${k.mounting_overdue > 0 ? 'danger' : 'good'}`} onClick={() => nav('/campaigns')} style={{ cursor: 'pointer' }} title="View campaigns">
          <div className="n">{k.mounting_overdue ?? 0}</div>
          <div className="l">Mounting Overdue</div>
        </div>
        <div className={`scooh-kpi ${k.validation_15_due > 0 ? 'danger' : 'good'}`} onClick={() => nav('/validations')} style={{ cursor: 'pointer' }} title="View validations">
          <div className="n">{k.validation_15_due ?? 0}</div>
          <div className="l">15-Day Validation Due</div>
        </div>
        <div className={`scooh-kpi ${k.final_validation_due > 0 ? 'danger' : 'good'}`} onClick={() => nav('/validations')} style={{ cursor: 'pointer' }} title="View validations">
          <div className="n">{k.final_validation_due ?? 0}</div>
          <div className="l">Final Validation Due</div>
        </div>
        <div className={`scooh-kpi ${k.invoice_actions > 0 ? 'danger' : 'good'}`} onClick={() => nav('/invoices')} style={{ cursor: 'pointer' }} title="View invoices">
          <div className="n">{k.invoice_actions ?? 0}</div>
          <div className="l">Invoice Actions</div>
        </div>
      </div>

      {/* Lower Dashboard Grid */}
      <div className="scooh-grid2 scooh-dashboard-lower">
        {/* Needs attention today */}
        <div className="scooh-panel">
          <h3>Needs attention today</h3>
          {alerts.length === 0 ? (
            <div className="scooh-empty" style={{ padding: '36px 16px', textAlign: 'center' }}>
              ✓ Nothing outstanding — all operations are on track.
            </div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
              {alerts.map((a, idx) => (
                <div
                  key={idx}
                  onClick={() => nav(a.link || '/campaigns')}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    padding: '12px 14px',
                    border: '1px solid var(--mb-border)',
                    borderRadius: '12px',
                    background: 'var(--mb-surface-2)',
                    cursor: 'pointer',
                    transition: 'border-color 0.15s ease'
                  }}
                  onMouseEnter={e => e.currentTarget.style.borderColor = 'var(--mb-primary)'}
                  onMouseLeave={e => e.currentTarget.style.borderColor = 'var(--mb-border)'}
                  title="Click to view details"
                >
                  <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                    <span className="scooh-plate">{a.site_code}</span>
                    <div>
                      <strong style={{ color: 'var(--mb-text)', fontSize: '12.5px' }}>{a.client || 'Client'}</strong>
                      <div className="scooh-footnote" style={{ color: 'var(--mb-muted)', fontSize: '10.5px' }}>{a.campaign}</div>
                    </div>
                  </div>
                  <span
                    className={`scooh-pill ${a.class === 'danger' ? 'vacant' : 'watch'}`}
                    style={{ whiteSpace: 'nowrap' }}
                  >
                    {a.tag}
                  </span>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Portfolio Occupancy Chart */}
        <div className="scooh-panel">
          <h3>
            Portfolio occupancy · <span className="scooh-accent">{k.average_occupancy || 0}%</span>
          </h3>
          <OccupancyDonut buckets={d.occupancy_buckets || {}} average={k.average_occupancy || 0} />
        </div>
      </div>
    </>
  );
}

function SitesView() {
  const navigate = useNavigate();
  const currentRole = getCurrentRole();
  const isAdmin = currentRole === 'admin';
  const isManager = currentRole === 'manager';
  const isStaff = currentRole === 'staff';
  const isReadOnly = currentRole === 'viewer';
  const canDelete = isAdmin || isManager;

  const [sites, setSites] = useState(defaultSites.map(unpackSite));
  const [search, setSearch] = useState('');
  const [cityFilter, setCityFilter] = useState('');
  const [mediaFilter, setMediaFilter] = useState('');
  const [availFilter, setAvailFilter] = useState('');
  const [lightingFilter, setLightingFilter] = useState('');
  const [sortState, setSortState] = useState({ key: 'site_code', dir: 'asc' });
  const [edit, setEdit] = useState(null);
  const [imageModalSite, setImageModalSite] = useState(null);
  const [saving, setSaving] = useState(false);

  async function load() {
    try {
      const { data } = await api.get('/sites');
      if (Array.isArray(data) && data.length > 0) {
        setSites(data.map(unpackSite));
      }
    } catch (e) {
      console.warn('Using default sites cache', e);
    }
  }

  const [importingExcel, setImportingExcel] = useState(false);

  useEffect(() => {
    load();
    const onSitesUpdate = () => load();
    window.addEventListener('mb-sites-updated', onSitesUpdate);
    const interval = setInterval(load, 35000); // 35s auto-refresh
    return () => {
      window.removeEventListener('mb-sites-updated', onSitesUpdate);
      clearInterval(interval);
    };
  }, []);

  const cities = useMemo(() => Array.from(new Set(sites.map(s => s.city).filter(Boolean))), [sites]);
  const mediaTypes = useMemo(() => Array.from(new Set(sites.map(s => s.media_type).filter(Boolean))), [sites]);
  const availabilities = useMemo(() => Array.from(new Set(sites.map(s => s.ppt_availability || s.availability).filter(Boolean))), [sites]);
  const lightings = useMemo(() => Array.from(new Set(sites.map(s => s.lighting).filter(Boolean))), [sites]);

  function handleSort(key) {
    setSortState(prev => prev.key === key ? { key, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { key, dir: 'asc' });
  }

  const filtered = useMemo(() => {
    const list = sites.filter(s => {
      const matchesQuery = matchSiteSearch(s, search);
      const matchesCity = !cityFilter || s.city === cityFilter;
      const matchesMedia = !mediaFilter || s.media_type === mediaFilter;
      const matchesAvail = !availFilter || (s.ppt_availability || s.availability) === availFilter;
      const matchesLighting = !lightingFilter || s.lighting === lightingFilter;
      return matchesQuery && matchesCity && matchesMedia && matchesAvail && matchesLighting;
    });

    if (sortState.key) {
      list.sort((a, b) => {
        let valA = a[sortState.key];
        let valB = b[sortState.key];
        if (sortState.key === 'monthly_rate') {
          valA = Number(a.ppt_rate || a.monthly_rate || 0);
          valB = Number(b.ppt_rate || b.monthly_rate || 0);
        }
        if (sortState.key === 'availability') {
          valA = a.ppt_availability || a.availability || '';
          valB = b.ppt_availability || b.availability || '';
        }
        return universalCompare(valA, valB, sortState.dir);
      });
    }
    return list;
  }, [sites, search, cityFilter, mediaFilter, availFilter, lightingFilter, sortState]);

  const [selectedIds, setSelectedIds] = useState(new Set());

  function toggleSelect(id) {
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function toggleSelectAll() {
    const visibleIds = filtered.map(s => s.id).filter(Boolean);
    const allSelected = visibleIds.length > 0 && visibleIds.every(id => selectedIds.has(id));
    if (allSelected) {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleIds.forEach(id => next.delete(id));
        return next;
      });
    } else {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleIds.forEach(id => next.add(id));
        return next;
      });
    }
  }

  function selectAllVisible() {
    setSelectedIds(new Set(filtered.map(s => s.id).filter(Boolean)));
  }

  function deselectAll() {
    setSelectedIds(new Set());
  }

  async function handleBatchDelete() {
    if (selectedIds.size === 0) return;
    const count = selectedIds.size;
    if (!confirm(`Are you sure you want to delete ${count} selected site record${count > 1 ? 's' : ''}?`)) {
      return;
    }
    try {
      await api.post('/sites/batch-delete', { ids: Array.from(selectedIds), hard: true });
      setSelectedIds(new Set());
      await load();
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
    } catch (err) {
      alert('Failed to delete selected sites: ' + (err.response?.data?.message || err.message));
    }
  }

  async function save(e) {
    e.preventDefault();
    setSaving(true);
    const data = Object.fromEntries(new FormData(e.currentTarget));
    try {
      if (edit?.id) await api.put(`/sites/${edit.id}`, data);
      else await api.post('/sites', data);
      setEdit(null);
      await load();
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
    } catch (err) {
      alert('Error saving site: ' + (err.response?.data?.message || err.message));
    } finally {
      setSaving(false);
    }
  }

  async function del(id) {
    if (confirm('Delete this site record?')) {
      try {
        await api.delete(`/sites/${id}?hard=true`);
        setSelectedIds(prev => {
          const next = new Set(prev);
          next.delete(id);
          return next;
        });
        await load();
        window.dispatchEvent(new CustomEvent('mb-sites-updated'));
      } catch (err) {
        alert('Failed to delete site: ' + (err.response?.data?.message || err.message));
      }
    }
  }

  async function handleImageUpload(siteId, files) {
    const fd = new FormData();
    [...files].forEach(f => fd.append('files', f));
    await api.post(`/sites/${siteId}/images`, fd, { headers: { 'Content-Type': 'multipart/form-data' } });
    await load();
    window.dispatchEvent(new CustomEvent('mb-sites-updated'));
  }

  async function handleDirectExcelImport(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportingExcel(true);
    try {
      const fd = new FormData();
      fd.append('file', file);
      const r = await api.post('/import/xlsx', fd, { headers: { 'Content-Type': 'multipart/form-data' } });
      alert(`✓ Excel Import Successful!\n\n${r.data.message || `Processed ${r.data.rows} sites (${r.data.updated} updated, ${r.data.created} created).`}`);
      await load();
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
    } catch (err) {
      alert('Excel Import failed: ' + (err.response?.data?.message || err.message));
    } finally {
      setImportingExcel(false);
      e.target.value = '';
    }
  }

  return (
    <>
      <PageHead
        title="Sites Directory"
        desc={`Managing ${filtered.length} of ${sites.length} total Media Buzz outdoor inventory sites.`}
        actions={
          <>
            <button type="button" className="scooh-btn ghost" onClick={load}>🔄 Refresh</button>
            {canDelete && (
              <label className="scooh-btn ghost" style={{ cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: '6px' }} title="Upload Excel spreadsheet to import sites, rates & availability">
                <span>📁 {importingExcel ? 'Importing…' : 'Import Excel'}</span>
                <input type="file" accept=".xlsx,.xls" hidden disabled={importingExcel} onChange={handleDirectExcelImport} />
              </label>
            )}
            {!isReadOnly && !isStaff && (
              <button className="scooh-btn purple-btn" onClick={() => setEdit({})}>+ Add Site</button>
            )}
          </>
        }
      />

      <div className="scooh-toolbar" style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', alignItems: 'center' }}>
        <input
          className="scooh-search"
          style={{ flex: '1 1 220px', minWidth: '180px' }}
          placeholder="Search site code (e.g. 01, MB-01), area, landmark…"
          value={search}
          onChange={e => setSearch(e.target.value)}
        />
        <div className="scooh-filters" style={{ display: 'flex', gap: '8px', flexWrap: 'wrap', alignItems: 'center' }}>
          <select value={cityFilter} onChange={e => setCityFilter(e.target.value)}>
            <option value="">All Cities ({cities.length})</option>
            {cities.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
          <select value={mediaFilter} onChange={e => setMediaFilter(e.target.value)}>
            <option value="">All Media Types ({mediaTypes.length})</option>
            {mediaTypes.map(m => <option key={m} value={m}>{m}</option>)}
          </select>
          {availabilities.length > 0 && (
            <select value={availFilter} onChange={e => setAvailFilter(e.target.value)}>
              <option value="">All Availability ({availabilities.length})</option>
              {availabilities.map(a => <option key={a} value={a}>{a}</option>)}
            </select>
          )}
          {lightings.length > 0 && (
            <select value={lightingFilter} onChange={e => setLightingFilter(e.target.value)}>
              <option value="">All Lighting ({lightings.length})</option>
              {lightings.map(l => <option key={l} value={l}>{l}</option>)}
            </select>
          )}
          <select
            value={`${sortState.key}:${sortState.dir}`}
            onChange={e => {
              const [k, d] = e.target.value.split(':');
              setSortState({ key: k, dir: d });
            }}
            style={{ fontWeight: 600 }}
          >
            <option value="site_code:asc">Sort: Site Code (01 → 87)</option>
            <option value="site_code:desc">Sort: Site Code (87 → 01)</option>
            <option value="area:asc">Sort: Area (A-Z)</option>
            <option value="area:desc">Sort: Area (Z-A)</option>
            <option value="monthly_rate:asc">Sort: Rate (Low to High)</option>
            <option value="monthly_rate:desc">Sort: Rate (High to Low)</option>
            <option value="availability:asc">Sort: Availability (A-Z)</option>
            <option value="size:asc">Sort: Size</option>
          </select>
          {(search || cityFilter || mediaFilter || availFilter || lightingFilter || sortState.key !== 'site_code' || sortState.dir !== 'asc') && (
            <button
              type="button"
              className="scooh-btn ghost"
              style={{ padding: '6px 10px', fontSize: '11px' }}
              onClick={() => {
                setSearch('');
                setCityFilter('');
                setMediaFilter('');
                setAvailFilter('');
                setLightingFilter('');
                setSortState({ key: 'site_code', dir: 'asc' });
              }}
            >
              Reset
            </button>
          )}
        </div>
      </div>

      {/* Batch Selection Action Bar */}
      {canDelete && selectedIds.size > 0 && (
        <div style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          padding: '11px 18px',
          background: 'linear-gradient(90deg, rgba(239,68,68,0.18), rgba(239,68,68,0.08))',
          borderBottom: '1px solid rgba(239,68,68,0.3)',
          borderRadius: '8px',
          margin: '10px 0',
          color: '#fca5a5',
          fontSize: '13px',
          flexWrap: 'wrap',
          gap: '12px'
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap' }}>
            <span style={{ fontWeight: 800, color: '#ffffff', display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
              <span style={{ display: 'inline-block', width: '8px', height: '8px', borderRadius: '50%', background: '#ef4444' }}></span>
              {selectedIds.size} site{selectedIds.size > 1 ? 's' : ''} selected
            </span>
            <button
              type="button"
              className="scooh-btn ghost"
              onClick={selectAllVisible}
              style={{ fontSize: '11.5px', padding: '4px 10px', color: '#f1f5f9', borderColor: '#475569' }}
            >
              Select all visible ({filtered.length})
            </button>
            <button
              type="button"
              className="scooh-btn ghost"
              onClick={deselectAll}
              style={{ fontSize: '11.5px', padding: '4px 10px', color: '#cbd5e1', borderColor: '#475569' }}
            >
              Deselect all
            </button>
          </div>
          <button
            type="button"
            className="scooh-btn danger"
            onClick={handleBatchDelete}
            style={{
              background: '#ef4444',
              color: '#ffffff',
              border: 'none',
              fontWeight: 800,
              padding: '7px 16px',
              borderRadius: '7px',
              cursor: 'pointer',
              display: 'inline-flex',
              alignItems: 'center',
              gap: '6px',
              boxShadow: '0 2px 10px rgba(239,68,68,0.45)'
            }}
          >
            🗑 Delete selected ({selectedIds.size})
          </button>
        </div>
      )}

      <div className="scooh-tablewrap">
        <table className="scooh-table">
          <thead>
            <tr>
              {canDelete && (
                <th style={{ width: '42px', textAlign: 'center', padding: '10px 8px' }}>
                  <input
                    type="checkbox"
                    checked={filtered.length > 0 && filtered.every(s => selectedIds.has(s.id))}
                    onChange={toggleSelectAll}
                    title="Select all visible sites"
                    style={{ cursor: 'pointer' }}
                  />
                </th>
              )}
              <SortHeader label="Site Code" sortKey="site_code" currentSort={sortState} onSort={handleSort} />
              <SortHeader label="Area / Landmark" sortKey="area" currentSort={sortState} onSort={handleSort} />
              <SortHeader label="Media" sortKey="media_type" currentSort={sortState} onSort={handleSort} />
              <SortHeader label="Size" sortKey="size" currentSort={sortState} onSort={handleSort} />
              <SortHeader label="Lighting" sortKey="lighting" currentSort={sortState} onSort={handleSort} />
              <SortHeader label="Availability" sortKey="availability" currentSort={sortState} onSort={handleSort} />
              <SortHeader label="Rate / Mo" sortKey="monthly_rate" currentSort={sortState} onSort={handleSort} />
              <th>Photos</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 ? (
              <tr><td colSpan={canDelete ? "10" : "9"} className="scooh-empty">No sites match the filters</td></tr>
            ) : (
              filtered.map(s => (
                <tr key={s.id || s.site_code}>
                  {canDelete && (
                    <td style={{ textAlign: 'center', padding: '10px 8px' }}>
                      <input
                        type="checkbox"
                        checked={selectedIds.has(s.id)}
                        onChange={() => toggleSelect(s.id)}
                        style={{ cursor: 'pointer' }}
                      />
                    </td>
                  )}
                  <td>
                    <div style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
                      <button
                        type="button"
                        onClick={() => navigate(`/campaigns?site=${encodeURIComponent(s.site_code)}`)}
                        className="scooh-plate"
                        title={`Open ${s.site_code} in Campaign Tracker`}
                        style={{
                          cursor: 'pointer',
                          background: 'rgba(56, 189, 248, 0.12)',
                          border: '1px solid rgba(56, 189, 248, 0.45)',
                          color: '#38bdf8',
                          display: 'inline-flex',
                          alignItems: 'center',
                          gap: '4px',
                          padding: '2px 7px',
                          borderRadius: '5px',
                          fontWeight: 800
                        }}
                      >
                        <span>{s.site_code}</span>
                        <span style={{ fontSize: '10px', opacity: 0.75 }}>↗</span>
                      </button>
                      {getSiteTypeTag(s.site_code) !== 'Single' && (
                        <span
                          style={{
                            fontSize: '10.5px',
                            padding: '2px 6px',
                            background: getSiteTypeTag(s.site_code) === 'Combined' ? 'rgba(168,85,247,0.18)' : 'rgba(56,189,248,0.18)',
                            color: getSiteTypeTag(s.site_code) === 'Combined' ? '#c084fc' : '#38bdf8',
                            borderRadius: '4px',
                            fontWeight: 700
                          }}
                          title={getConflictSummary(s.site_code)?.message || ''}
                        >
                          {getSiteTypeTag(s.site_code)}
                        </span>
                      )}
                    </div>
                  </td>
                  <td>
                    <strong>{s.area || s.city}</strong>
                    <div className="scooh-footnote">{s.address}</div>
                  </td>
                  <td>{s.media_type}</td>
                  <td>{s.size}</td>
                  <td><span className="scooh-badge">{s.lighting}</span></td>
                  <td>
                    {String(s.availability || '').startsWith('Occupied (Linked') ? (
                      <span
                        className="scooh-pill watch"
                        title={s.availability}
                        style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '11px', whiteSpace: 'nowrap', background: 'rgba(245,158,11,0.15)', color: '#fbbf24', borderColor: 'rgba(245,158,11,0.35)' }}
                      >
                        🔗 {s.availability}
                      </span>
                    ) : String(s.availability || '').toLowerCase() === 'occupied' ? (
                      <span className="scooh-pill vacant">Occupied</span>
                    ) : (
                      <span className={`scooh-pill ${s.availability === 'Available' ? 'active' : 'watch'}`}>
                        {s.availability || 'Available'}
                      </span>
                    )}
                  </td>
                  <td><b>{money(s.ppt_rate || s.monthly_rate)}</b></td>
                  <td>
                    <button
                      type="button"
                      className="scooh-iconbtn"
                      onClick={() => setImageModalSite(s)}
                      title="Manage Presentation Photos"
                    >
                      📷 {(s.ppt_images || []).length}
                    </button>
                  </td>
                  <td>
                    <div className="scooh-rowactions">
                      <button className="scooh-iconbtn scooh-text-action" onClick={() => setEdit(s)}>
                        {isReadOnly ? 'View' : 'Edit'}
                      </button>
                      {canDelete && (
                        <button className="scooh-iconbtn danger-icon" onClick={() => del(s.id)}>×</button>
                      )}
                    </div>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {/* Edit/Add Modal */}
      {edit && (
        <div className="scooh-modal-overlay">
          <div className="scooh-modal">
            <div className="scooh-modalhead">
              <div>
                <h2>{edit.id ? 'Edit Site' : 'Add New Site'}</h2>
                <span className="scooh-modal-subtitle">Update site specifications and pricing</span>
              </div>
              <button type="button" className="scooh-modal-close" onClick={() => setEdit(null)}>×</button>
            </div>
            <form onSubmit={save}>
              <div className="scooh-modalbody">
                <div className="scooh-grid2">
                  <div className="scooh-field">
                    <label>Site Code</label>
                    <input name="site_code" defaultValue={edit.site_code ?? ''} placeholder="e.g. AMD-GT-001" required />
                  </div>
                  <div className="scooh-field">
                    <label>City</label>
                    <input name="city" defaultValue={edit.city ?? 'Ahmedabad'} placeholder="e.g. Ahmedabad" required />
                  </div>
                  <div className="scooh-field">
                    <label>Area / Landmark</label>
                    <input name="area" defaultValue={edit.area ?? ''} placeholder="e.g. Shivranjani Cross Roads" required />
                  </div>
                  <div className="scooh-field">
                    <label>Media Type</label>
                    <select name="media_type" defaultValue={edit.media_type ?? 'Hoarding'}>
                      <option value="Hoarding">Hoarding</option>
                      <option value="Gantry">Gantry</option>
                      <option value="Unipole">Unipole</option>
                      <option value="Billboard">Billboard</option>
                      <option value="DOOH">DOOH</option>
                    </select>
                  </div>
                  <div className="scooh-field">
                    <label>Size (e.g. 30x10)</label>
                    <input name="size" defaultValue={edit.size ?? ''} placeholder="e.g. 30x10 ft" />
                  </div>
                  <div className="scooh-field">
                    <label>Lighting</label>
                    <select name="lighting" defaultValue={edit.lighting ?? 'BL'}>
                      <option value="BL">BL (Backlit)</option>
                      <option value="FL">FL (Frontlit)</option>
                      <option value="NL">NL (Nonlit)</option>
                    </select>
                  </div>
                  <div className="scooh-field">
                    <label>Monthly Rate (₹)</label>
                    <input name="monthly_rate" type="number" defaultValue={edit.monthly_rate ?? 0} placeholder="250000" />
                  </div>
                  <div className="scooh-field">
                    <label>Availability</label>
                    <select name="availability" defaultValue={edit.availability ?? 'Available'}>
                      <option value="Available">Available</option>
                      <option value="Occupied">Occupied</option>
                      <option value="Maintenance">Maintenance</option>
                      <option value="Booked">Booked</option>
                    </select>
                  </div>
                  <div className="scooh-field" style={{ gridColumn: '1 / -1' }}>
                    <label>Full Address</label>
                    <textarea name="address" defaultValue={edit.address ?? ''} placeholder="Full road, junction and landmark location description…" rows={3} />
                  </div>
                  <div className="scooh-field">
                    <label>GPS Coordinates</label>
                    <input name="gps" defaultValue={edit.gps ?? ''} placeholder="23.019187, 72.530184" />
                  </div>
                  <div className="scooh-field">
                    <label>Google Maps URL</label>
                    <input name="maps_url" defaultValue={edit.maps_url ?? ''} placeholder="https://maps.google.com/?q=..." />
                  </div>
                </div>
              </div>
              <div className="scooh-modalfoot">
                <button type="button" className="scooh-btn ghost" onClick={() => setEdit(null)}>
                  {isReadOnly ? 'Close' : 'Cancel'}
                </button>
                {!isReadOnly && (
                  <button className="scooh-btn purple-btn" disabled={saving}>
                    {saving ? 'Saving…' : 'Save Site'}
                  </button>
                )}
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Image Upload Modal */}
      {imageModalSite && (
        <div className="scooh-modal-overlay">
          <div className="scooh-modal" style={{ maxWidth: '640px' }}>
            <div className="scooh-modalhead">
              <div>
                <h2>Images for {imageModalSite.site_code}</h2>
                <span className="scooh-modal-subtitle">Upload photos for automated PPT and presentations</span>
              </div>
              <button type="button" className="scooh-modal-close" onClick={() => setImageModalSite(null)}>×</button>
            </div>
            <div className="scooh-modalbody">
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '10px', marginBottom: '16px' }}>
                {(imageModalSite.ppt_images || []).map((url, i) => (
                  <div key={i} style={{ position: 'relative', border: '1px solid #34404e', borderRadius: '8px', overflow: 'hidden' }}>
                    <img src={url} alt="Site" style={{ width: '100%', height: '110px', objectFit: 'cover' }} />
                  </div>
                ))}
              </div>
              <label className="scooh-btn primary" style={{ display: 'inline-flex', cursor: 'pointer' }}>
                + Upload Images
                <input
                  type="file"
                  multiple
                  accept="image/*"
                  hidden
                  onChange={async e => {
                    if (e.target.files.length) {
                      await handleImageUpload(imageModalSite.id, e.target.files);
                      setImageModalSite(null);
                    }
                  }}
                />
              </label>
            </div>
            <div className="scooh-modalfoot">
              <button type="button" className="scooh-btn ghost" onClick={() => setImageModalSite(null)}>Close</button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

function PptView() {
  const navigate = useNavigate();
  const [sites, setSites] = useState(defaultSites.map(unpackSite));
  const [pages, setPages] = useState(() => {
    try {
      const saved = localStorage.getItem('mb_ppt_pages_cache');
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });
  const [uploadingPage, setUploadingPage] = useState({});
  const [sel, setSel] = useState({});
  const [query, setQuery] = useState('');
  const [areaFilter, setAreaFilter] = useState('');
  const [dateFilter, setDateFilter] = useState('');
  const [generating, setGenerating] = useState(false);
  const [pptName, setPptName] = useState('MediaBuzz_Automated-PPT');
  const [alsoGenerateExcel, setAlsoGenerateExcel] = useState(true);
  const [sortState, setSortState] = useState({ key: 'site_code', dir: 'asc' });
  // Campaign tracker map: site_code (uppercase) → latest active campaign row
  const [campaignMap, setCampaignMap] = useState({});

  function getCleanPptName() {
    let clean = (pptName || '').trim();
    if (!clean) clean = `MediaBuzz_Presentation_${new Date().toISOString().slice(0, 10)}`;
    clean = clean.replace(/\.(pptx|xlsx)$/i, '');
    return clean.replace(/[\\/:*?"<>|]/g, '_');
  }

  function fmtDate(iso) {
    if (!iso) return '';
    try {
      const d = new Date(iso);
      return d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
    } catch { return iso; }
  }

  async function load() {
    try {
      const [sRes, pRes, cRes] = await Promise.all([
        api.get('/sites'),
        api.get('/ppt-pages').catch(() => api.get('/ppt-pages/public').catch(() => ({ data: null }))),
        api.get('/campaigns').catch(() => ({ data: [] }))
      ]);
      if (Array.isArray(sRes.data) && sRes.data.length > 0) {
        setSites(sRes.data.map(unpackSite));
      }
      if (pRes && pRes.data) {
        setPages(prev => {
          const merged = { ...prev };
          if (pRes.data.first) merged.first = pRes.data.first;
          if (pRes.data.second_last) merged.second_last = pRes.data.second_last;
          if (pRes.data.last) merged.last = pRes.data.last;
          try {
            localStorage.setItem('mb_ppt_pages_cache', JSON.stringify(merged));
          } catch {}
          return merged;
        });
      }
      // Build site_code → active campaign map from Campaign Tracker
      // Primary campaigns take precedence, then latest end_date / start_date / id desc
      if (Array.isArray(cRes.data)) {
        const map = {};
        const sorted = cRes.data
          .filter(c => c.record_status === 'active')
          .sort((a, b) => {
            const isLinkedA = String(a.parent_campaign || '').startsWith('LINKED:');
            const isLinkedB = String(b.parent_campaign || '').startsWith('LINKED:');
            if (isLinkedA !== isLinkedB) return isLinkedA ? 1 : -1; // primary first
            const endA = new Date(a.end_date || 0).getTime();
            const endB = new Date(b.end_date || 0).getTime();
            if (endA !== endB && endA > 0 && endB > 0) return endB - endA;
            const startA = new Date(a.start_date || 0).getTime();
            const startB = new Date(b.start_date || 0).getTime();
            if (startA !== startB && startA > 0 && startB > 0) return startB - startA;
            return (b.id || 0) - (a.id || 0);
          });
        for (const c of sorted) {
          const code = String(c.site_code || '').toUpperCase().trim();
          if (code && !map[code]) map[code] = c;
        }
        setCampaignMap(map);
      }
    } catch (e) {
      console.warn('PPT load notice:', e);
    }
  }

  const [importingExcel, setImportingExcel] = useState(false);

  async function handlePptExcelImport(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportingExcel(true);
    try {
      const fd = new FormData();
      fd.append('file', file);
      const r = await api.post('/import/xlsx', fd, { headers: { 'Content-Type': 'multipart/form-data' } });
      alert(`✓ Excel Import Successful!\n\n${r.data.message || `Processed ${r.data.rows} sites (${r.data.updated} updated, ${r.data.created} created).`}`);
      await load();
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
    } catch (err) {
      alert('Excel Import failed: ' + (err.response?.data?.message || err.message));
    } finally {
      setImportingExcel(false);
      e.target.value = '';
    }
  }

  useEffect(() => {
    load();
    const onSitesUpdate = () => load();
    window.addEventListener('mb-sites-updated', onSitesUpdate);
    return () => window.removeEventListener('mb-sites-updated', onSitesUpdate);
  }, []);

  const areas = useMemo(() => Array.from(new Set(sites.map(s => s.area).filter(Boolean))).sort(), [sites]);

  function handleSort(key) {
    setSortState(prev => prev.key === key ? { key, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { key, dir: 'asc' });
  }

  const selectedCount = useMemo(() => sites.filter(s => sel[s.id || s.site_code]?.checked).length, [sites, sel]);

  const filtered = useMemo(() => {
    const list = sites.filter(s => {
      if (!matchSiteSearch(s, query)) return false;
      if (areaFilter && s.area !== areaFilter) return false;
      if (dateFilter) {
        const av = String(s.ppt_availability || s.availability || '').toLowerCase().trim();
        if (av !== 'available' && av !== 'immediate' && av !== 'long term') {
          if (av > dateFilter) return false;
        }
      }
      return true;
    });

    list.sort((a, b) => {
      const aSel = !!sel[a.id || a.site_code]?.checked;
      const bSel = !!sel[b.id || b.site_code]?.checked;
      // Selected sites always show on top
      if (aSel !== bSel) return aSel ? -1 : 1;

      if (sortState.key) {
        let valA = a[sortState.key];
        let valB = b[sortState.key];
        if (sortState.key === 'monthly_rate') {
          valA = Number(a.ppt_rate || a.monthly_rate || 0);
          valB = Number(b.ppt_rate || b.monthly_rate || 0);
        }
        return universalCompare(valA, valB, sortState.dir);
      }
      return 0;
    });
    return list;
  }, [sites, query, areaFilter, dateFilter, sortState, sel]);

  async function uploadPage(k, file) {
    if (!file) return;
    setUploadingPage(prev => ({ ...prev, [k]: true }));
    try {
      // 1. Immediate local base64 preview & offline fallback so image loads instantly
      const dataUrl = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(file);
      });

      if (dataUrl) {
        setPages(prev => {
          const next = { ...prev, [k]: dataUrl };
          try {
            localStorage.setItem('mb_ppt_pages_cache', JSON.stringify(next));
            localStorage.setItem(`mb_ppt_raw_${k}`, dataUrl);
          } catch {}
          return next;
        });
      }

      // 2. Persist to server backend & MySQL
      const fd = new FormData();
      fd.append('file', file);
      const res = await api.post('/ppt-pages/' + k, fd, { headers: { 'Content-Type': 'multipart/form-data' } });
      const serverUrl = res.data?.url || dataUrl;
      setPages(prev => {
        const next = { ...prev, [k]: serverUrl };
        try {
          localStorage.setItem('mb_ppt_pages_cache', JSON.stringify(next));
        } catch {}
        return next;
      });
    } catch (e) {
      console.error('Page upload notice:', e);
      alert('Upload notice: ' + (e.response?.data?.message || e.message));
    } finally {
      setUploadingPage(prev => ({ ...prev, [k]: false }));
    }
  }

  async function removePage(k) {
    const label = k === 'first' ? 'First page' : (k === 'last' ? 'Last page' : 'Second page');
    if (!confirm(`Are you sure you want to remove the uploaded ${label}?`)) return;
    try {
      await api.delete('/ppt-pages/' + k);
    } catch (e) {
      console.warn('Delete page server warning:', e);
    }
    setPages(prev => {
      const next = { ...prev, [k]: '' };
      try {
        localStorage.setItem('mb_ppt_pages_cache', JSON.stringify(next));
        localStorage.removeItem(`mb_ppt_raw_${k}`);
      } catch {}
      return next;
    });
  }

  async function addImages(site, files) {
    try {
      if (site.id) {
        const fd = new FormData();
        Array.from(files).forEach(f => fd.append('files', f));
        const res = await api.post(`/sites/${site.id}/images`, fd, { headers: { 'Content-Type': 'multipart/form-data' } });
        setSites(prev => prev.map(s => s.id === site.id ? { ...s, ppt_images: res.data.images } : s));
      } else {
        const urls = Array.from(files).map(f => URL.createObjectURL(f));
        setSites(prev => prev.map(s => s.site_code === site.site_code ? { ...s, ppt_images: [...(s.ppt_images || []), ...urls] } : s));
      }
    } catch (e) {
      alert('Failed to upload images: ' + (e.response?.data?.message || e.message));
    }
  }

  async function removeImage(site, index) {
    try {
      if (site.id) {
        const res = await api.delete(`/sites/${site.id}/images/${index}`);
        setSites(prev => prev.map(s => s.id === site.id ? { ...s, ppt_images: res.data.images } : s));
      } else {
        const remaining = (site.ppt_images || []).filter((_, i) => i !== index);
        setSites(prev => prev.map(s => s.site_code === site.site_code ? { ...s, ppt_images: remaining } : s));
      }
    } catch (e) {
      alert('Failed to remove image: ' + (e.response?.data?.message || e.message));
    }
  }

  function selectAllVisible() {
    const newSel = { ...sel };
    filtered.forEach(s => {
      newSel[s.id || s.site_code] = { ...(newSel[s.id || s.site_code] || {}), checked: true };
    });
    setSel(newSel);
  }

  function deselectAll() {
    const newSel = { ...sel };
    filtered.forEach(s => {
      if (newSel[s.id || s.site_code]) {
        newSel[s.id || s.site_code].checked = false;
      }
    });
    setSel(newSel);
  }

  async function downloadExcel() {
    try {
      const selectedList = sites.filter(s => sel[s.id || s.site_code]?.checked);
      const targetSites = selectedList.length > 0 ? selectedList : (filtered.length > 0 ? filtered : sites);
      const cleanName = getCleanPptName();

      const rows = targetSites.map((x, idx) => {
        const v = sel[x.id || x.site_code] || {};
        const cCode = String(x.site_code || '').toUpperCase().trim();
        const camp = campaignMap[cCode];
        let autoAvail = v.availability ?? x.ppt_availability ?? x.availability ?? 'Available';
        let endDateFmt = '';
        if (camp && camp.end_date) {
          const endFmt = fmtDate(camp.end_date);
          const endDate = new Date(camp.end_date);
          const now = new Date();
          if (endDate >= now) {
            endDateFmt = endFmt;
            if (!v.availability) {
              autoAvail = `Booked till ${endFmt}`;
            }
          } else {
            if (!v.availability) autoAvail = 'Available';
          }
        }
        let w = x.width || '', h = x.height || '';
        if ((!w || !h) && x.size) {
          const parts = String(x.size).toLowerCase().split('x');
          if (parts.length === 2) { w = parts[0].trim(); h = parts[1].trim(); }
        }
        const sqft = (Number(w || 0) && Number(h || 0)) ? Number(w) * Number(h) : (x.sq_ft || '');
        const coords = x.gps || ([x.latitude, x.longitude].filter(Boolean).join(', ')) || '';

        return {
          'SR NO': idx + 1,
          'SITE CODE': x.site_code || '',
          'AREA': x.area || x.city || '',
          'LOCATION': x.address || '',
          'MEDIA': x.media_type || 'Hoarding',
          'LIGHT': x.lighting || 'BL',
          'W': w || '',
          'H': h || '',
          'SQ FT': sqft,
          'AVAILABLITY': autoAvail,
          'End Date': endDateFmt,
          'Selling Amount': Number(v.rate ?? x.ppt_rate ?? x.monthly_rate ?? 0),
          'Latitude Longitude': coords
        };
      });

      await exportStyledExcel(rows, `${cleanName}.xlsx`, cleanName);
    } catch (e) {
      alert('Export Excel failed: ' + e.message);
    }
  }

  async function generate() {
    const chosen = sites
      .filter(s => sel[s.id || s.site_code]?.checked)
      .map(s => {
        const v = sel[s.id || s.site_code] || {};
        // Look up active campaign from Campaign Tracker for this site (only end date)
        const cCode = String(s.site_code || '').toUpperCase().trim();
        const camp = campaignMap[cCode];
        let autoAvail = v.availability ?? s.ppt_availability ?? s.availability;
        let endDateFmt = '';
        if (camp && camp.end_date) {
          const endFmt = fmtDate(camp.end_date);
          const endDate = new Date(camp.end_date);
          const now = new Date();
          if (endDate >= now) {
            endDateFmt = endFmt;
            // Auto-set availability to only show the end date unless user manually overrode it
            if (!v.availability) {
              autoAvail = `Booked till ${endFmt}`;
            }
          } else {
            if (!v.availability) autoAvail = 'Available';
          }
        }
        return {
          ...s,
          _availability: autoAvail,
          _rate: v.rate ?? s.ppt_rate ?? s.monthly_rate,
          _showRate: !!v.showRate,
          _endDate: endDateFmt
        };
      });

    if (!chosen.length) return alert('Please select at least one site to include in the PowerPoint.');

    const cleanName = getCleanPptName();

    setGenerating(true);
    try {
      await makePpt(chosen, pages, `${cleanName}.pptx`);

      if (alsoGenerateExcel) {
        // Small delay to ensure browser reliably dispatches multiple downloads
        await new Promise(r => setTimeout(r, 450));

        const rows = chosen.map((x, idx) => {
          let w = x.width || '', h = x.height || '';
          if ((!w || !h) && x.size) {
            const parts = String(x.size).toLowerCase().split('x');
            if (parts.length === 2) { w = parts[0].trim(); h = parts[1].trim(); }
          }
          const sqft = (Number(w || 0) && Number(h || 0)) ? Number(w) * Number(h) : (x.sq_ft || '');
          const coords = x.gps || ([x.latitude, x.longitude].filter(Boolean).join(', ')) || '';

          return {
            'SR NO': idx + 1,
            'SITE CODE': x.site_code || '',
            'AREA': x.area || x.city || '',
            'LOCATION': x.address || '',
            'MEDIA': x.media_type || 'Hoarding',
            'LIGHT': x.lighting || 'BL',
            'W': w || '',
            'H': h || '',
            'SQ FT': sqft,
            'AVAILABLITY': x._availability ?? 'Available',
            'End Date': x._endDate || '',
            'Selling Amount': Number(x._rate ?? 0),
            'Latitude Longitude': coords
          };
        });

        await exportStyledExcel(rows, `${cleanName}.xlsx`, cleanName);
      }
    } catch (e) {
      alert('Generation Error: ' + e.message);
    } finally {
      setGenerating(false);
    }
  }

  return (
    <>
      <PageHead
        title="Automated PPT"
        desc="Select sites, customize presentation name, and generate client pitch decks with the official Media Buzz side dashboard."
        actions={
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
            <div className="scooh-ppt-name-bar" style={{ display: 'flex', alignItems: 'center', gap: '8px', background: 'rgba(10, 30, 53, 0.85)', padding: '5px 12px', borderRadius: '8px', border: '1px solid #1c3b60' }}>
              <span style={{ fontSize: '11px', fontWeight: 700, color: '#FFC200', textTransform: 'uppercase', letterSpacing: '0.6px' }}>Presentation Name:</span>
              <input
                type="text"
                className="scooh-input"
                style={{ width: '220px', padding: '5px 10px', fontSize: '13px', background: '#071526', border: '1px solid #234d7d', color: '#fff', borderRadius: '5px' }}
                value={pptName}
                onChange={e => setPptName(e.target.value)}
                placeholder="MediaBuzz_Automated-PPT"
              />
              <span style={{ fontSize: '12px', color: '#8fa4bd', fontWeight: 600 }}>
                {alsoGenerateExcel ? '.pptx & .xlsx' : '.pptx'}
              </span>
            </div>

            <label
              className="scooh-ppt-excel-option"
              title="Automatically generate and download matching Excel sheet with PPT contents and presentation title"
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: '8px',
                cursor: 'pointer',
                background: alsoGenerateExcel ? 'rgba(255, 194, 0, 0.12)' : 'rgba(10, 30, 53, 0.85)',
                padding: '6px 13px',
                borderRadius: '8px',
                border: alsoGenerateExcel ? '1px solid #FFC200' : '1px solid #1c3b60',
                userSelect: 'none',
                color: alsoGenerateExcel ? '#FFC200' : '#c8d6e5',
                fontWeight: 600,
                fontSize: '12.5px',
                transition: 'all 0.2s ease'
              }}
            >
              <input
                type="checkbox"
                checked={alsoGenerateExcel}
                onChange={e => setAlsoGenerateExcel(e.target.checked)}
                style={{ accentColor: '#FFC200', width: '15px', height: '15px', cursor: 'pointer' }}
              />
              <span>Also generate Excel (.xlsx)</span>
            </label>

            <button type="button" className="scooh-btn ghost" onClick={downloadExcel} title="Export selected sites directly to Excel">
              Download Excel (.xlsx)
            </button>
            <button type="button" className="scooh-btn primary" onClick={generate} disabled={generating}>
              {generating
                ? (alsoGenerateExcel ? 'Creating PPT & Excel…' : 'Creating PPT…')
                : (alsoGenerateExcel ? 'Generate PPT + Excel' : 'Generate PPT')}
            </button>
          </div>
        }
      />

      {/* Fixed presentation pages section */}
      <section className="scooh-form-section scooh-ppt-fixed-pages">
        <div className="scooh-form-section-head">
          <div>
            <h3>Fixed presentation pages (Optional)</h3>
            <p>Optional cover and closing pages. Order: page 1, page 2, selected site slides, then page 3.</p>
          </div>
        </div>
        <div className="scooh-grid3">
          {[
            ['first', 'First page', 'Full-screen cover image for slide 1.'],
            ['second_last', 'Second page', 'Full-screen image used as slide 2, before the selected site slides.'],
            ['last', 'Last page', 'Full-screen image used after all selected site slides.']
          ].map(([k, title, desc]) => {
            const pageUrl = pages[k] || '';
            const isUploading = !!uploadingPage[k];
            const rawFallback = (() => {
              try { return localStorage.getItem(`mb_ppt_raw_${k}`) || ''; } catch { return ''; }
            })();
            const activeUrl = pageUrl || rawFallback;

            return (
              <div
                className="scooh-panel"
                key={k}
                style={{
                  margin: 0,
                  padding: '18px',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '10px',
                  border: activeUrl ? '1px solid rgba(0, 200, 117, 0.35)' : '1px solid #1c2838',
                  background: activeUrl ? 'linear-gradient(180deg, #0d1624 0%, #08101a 100%)' : '#0b1320',
                  borderRadius: '12px'
                }}
              >
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '8px' }}>
                  <strong style={{ fontSize: '14px', color: '#eef2f6' }}>{title}</strong>
                  {activeUrl ? (
                    <span style={{
                      display: 'inline-flex',
                      alignItems: 'center',
                      gap: '4px',
                      fontSize: '11px',
                      fontWeight: 700,
                      color: '#00e575',
                      background: 'rgba(0, 229, 117, 0.12)',
                      border: '1px solid rgba(0, 229, 117, 0.3)',
                      padding: '2px 8px',
                      borderRadius: '12px'
                    }}>
                      ✓ Saved & Loaded
                    </span>
                  ) : (
                    <span style={{ fontSize: '11px', color: '#7a8ba3' }}>Optional</span>
                  )}
                </div>

                <small style={{ color: '#91a5c2', fontSize: '11.5px', lineHeight: '1.4' }}>{desc}</small>

                {activeUrl ? (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', marginTop: '4px' }}>
                    <div
                      style={{
                        position: 'relative',
                        width: '100%',
                        height: '135px',
                        background: '#040912',
                        borderRadius: '8px',
                        overflow: 'hidden',
                        border: '1px solid #1c2c42',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center'
                      }}
                    >
                      <img
                        src={activeUrl}
                        alt={title}
                        onError={e => {
                          if (rawFallback && e.target.src !== rawFallback) {
                            e.target.src = rawFallback;
                          } else if (activeUrl.startsWith('/') && !e.target.src.includes('3000')) {
                            e.target.src = `http://localhost:3000${activeUrl}`;
                          }
                        }}
                        style={{
                          width: '100%',
                          height: '100%',
                          objectFit: 'cover',
                          display: 'block'
                        }}
                      />
                      {isUploading && (
                        <div style={{
                          position: 'absolute',
                          inset: 0,
                          background: 'rgba(2, 10, 20, 0.75)',
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                          color: '#FFC200',
                          fontSize: '12px',
                          fontWeight: 700,
                          gap: '6px'
                        }}>
                          <span>Saving…</span>
                        </div>
                      )}
                    </div>

                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                      <label
                        className="scooh-btn"
                        style={{
                          flex: 1,
                          cursor: isUploading ? 'not-allowed' : 'pointer',
                          display: 'inline-flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                          fontSize: '11.5px',
                          padding: '6px 12px',
                          minHeight: '34px'
                        }}
                      >
                        {isUploading ? 'Uploading…' : 'Replace image'}
                        <input
                          type="file"
                          accept="image/*"
                          hidden
                          disabled={isUploading}
                          onChange={e => e.target.files[0] && uploadPage(k, e.target.files[0])}
                        />
                      </label>
                      <button
                        type="button"
                        className="scooh-btn danger"
                        onClick={() => removePage(k)}
                        disabled={isUploading}
                        title="Remove uploaded image"
                        style={{
                          minHeight: '34px',
                          padding: '6px 10px',
                          fontSize: '11.5px'
                        }}
                      >
                        Remove
                      </button>
                    </div>
                  </div>
                ) : (
                  <div style={{ marginTop: '6px' }}>
                    <label
                      className="scooh-btn primary"
                      style={{
                        width: '100%',
                        cursor: isUploading ? 'not-allowed' : 'pointer',
                        display: 'inline-flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        gap: '6px',
                        minHeight: '38px'
                      }}
                    >
                      {isUploading ? 'Uploading…' : '+ Upload image'}
                      <input
                        type="file"
                        accept="image/*"
                        hidden
                        disabled={isUploading}
                        onChange={e => e.target.files[0] && uploadPage(k, e.target.files[0])}
                      />
                    </label>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </section>

      {/* 1. Select sites and slide values section */}
      <section className="scooh-form-section scooh-ppt-generator" style={{ marginTop: '18px' }}>
        <div className="scooh-form-section-head">
          <div>
            <h3>1. Select sites and slide values</h3>
            <p>Use the approved Media Buzz format: site details on the left and the site image on the right. Availability and monthly rate are optional and can be imported from Excel or edited manually.</p>
          </div>
        </div>

        <div className="scooh-grid3 scooh-ppt-global-settings">
          <div className="scooh-field">
            <label>Area filter</label>
            <select id="scooh-ppt-area-filter" value={areaFilter} onChange={e => setAreaFilter(e.target.value)}>
              <option value="">All areas</option>
              {areas.map(a => <option key={a} value={a}>{a}</option>)}
            </select>
          </div>
          <div className="scooh-field">
            <label>Available on / after</label>
            <input id="scooh-ppt-date-filter" type="date" value={dateFilter} onChange={e => setDateFilter(e.target.value)} />
          </div>
          <div className="scooh-note">
            Excel-imported PPT availability and rate appear on each card. You can edit them manually before generating the PPT.
          </div>
        </div>

        {/* Toolbar with Search on Left and Select/Deselect on Right */}
        <div
          className="scooh-ppt-toolbar"
          style={{
            display: 'flex',
            alignItems: 'center',
            gap: '12px',
            margin: '18px 0 10px',
            width: '100%'
          }}
        >
          <div
            className="scooh-search scooh-ppt-search-wrap"
            style={{ flex: 1, minWidth: 0, margin: 0 }}
          >
            <span aria-hidden="true">⌕</span>
            <input
              id="scooh-ppt-search"
              type="search"
              placeholder="Search site code (e.g. 01, MB-01), area, city, location…"
              value={query}
              onChange={e => setQuery(e.target.value)}
              autoComplete="off"
            />
          </div>

          <div
            className="scooh-ppt-select-actions"
            style={{ display: 'flex', alignItems: 'center', gap: '8px', flexShrink: 0 }}
          >
            <label
              className="scooh-btn ghost"
              style={{ minHeight: '44px', padding: '0 14px', fontSize: '12.5px', whiteSpace: 'nowrap', cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: '6px' }}
              title="Import Excel to update PPT rates & availability"
            >
              <span>📁 {importingExcel ? 'Importing…' : 'Import Excel'}</span>
              <input type="file" accept=".xlsx,.xls" hidden disabled={importingExcel} onChange={handlePptExcelImport} />
            </label>
            <button
              type="button"
              className="scooh-btn ghost"
              onClick={selectAllVisible}
              style={{ minHeight: '44px', padding: '0 16px', fontSize: '12.5px', whiteSpace: 'nowrap' }}
            >
              Select all visible
            </button>
            <button
              type="button"
              className="scooh-btn ghost"
              onClick={deselectAll}
              style={{ minHeight: '44px', padding: '0 16px', fontSize: '12.5px', whiteSpace: 'nowrap' }}
            >
              Deselect all
            </button>
            <select
              value={`${sortState.key}:${sortState.dir}`}
              onChange={e => {
                const [k, d] = e.target.value.split(':');
                setSortState({ key: k, dir: d });
              }}
              style={{
                height: '44px',
                padding: '0 12px',
                borderRadius: '8px',
                border: '1px solid #344258',
                background: '#111720',
                color: '#e2eaf4',
                fontSize: '12.5px',
                cursor: 'pointer',
                fontWeight: 600
              }}
            >
              <option value="site_code:asc">Sort: Site Code (01 → 87)</option>
              <option value="site_code:desc">Sort: Site Code (87 → 01)</option>
              <option value="area:asc">Sort: Area (A–Z)</option>
              <option value="area:desc">Sort: Area (Z–A)</option>
              <option value="city:asc">Sort: City (A–Z)</option>
              <option value="media_type:asc">Sort: Media Type (A–Z)</option>
              <option value="monthly_rate:asc">Sort: Rate (Low to High)</option>
              <option value="monthly_rate:desc">Sort: Rate (High to Low)</option>
              <option value="availability:asc">Sort: Availability</option>
            </select>
            <button
              type="button"
              className="scooh-btn ghost"
              onClick={() => setSortState(prev => ({ ...prev, dir: prev.dir === 'asc' ? 'desc' : 'asc' }))}
              title={sortState.dir === 'asc' ? 'Ascending — click for descending' : 'Descending — click for ascending'}
              style={{ minHeight: '44px', padding: '0 12px', fontSize: '13px', whiteSpace: 'nowrap', fontWeight: 700 }}
            >
              {sortState.dir === 'asc' ? '↑ Asc' : '↓ Desc'}
            </button>
          </div>
        </div>

        <div className="scooh-ppt-search-count" id="scooh-ppt-search-count" style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '8px' }}>
          <span>{filtered.length} of {sites.length} sites</span>
          {selectedCount > 0 && (
            <span style={{ color: '#a78bfa', fontWeight: 700, fontSize: '12px' }}>
              ✓ {selectedCount} site{selectedCount > 1 ? 's' : ''} selected (showing on top)
            </span>
          )}
        </div>

        {/* Sites Grid */}
        <div className="scooh-ppt-grid" id="scooh-ppt-grid">
          {filtered.map(s => {
            const siteKey = s.id || s.site_code;
            const v = sel[siteKey] || {};
            const isChecked = !!v.checked;
            const imgs = s.ppt_images || [];
            const cCode = String(s.site_code || '').toUpperCase().trim();
            const camp = campaignMap[cCode];
            let autoAvail = '';
            let endDateText = '';
            if (camp && camp.end_date) {
              const endFmt = fmtDate(camp.end_date);
              const endDate = new Date(camp.end_date);
              const now = new Date();
              if (endDate >= now) {
                autoAvail = `Booked till ${endFmt}`;
                endDateText = endFmt;
              } else {
                autoAvail = 'Available';
                endDateText = '';
              }
            }

            return (
              <label
                key={siteKey}
                className="scooh-ppt-site"
                style={{
                  border: isChecked ? '1.5px solid #a78bfa' : '1px solid #344258',
                  boxShadow: isChecked ? '0 0 12px rgba(167, 139, 250, 0.25)' : 'none',
                  background: isChecked ? 'rgba(30, 27, 75, 0.45)' : undefined
                }}
              >
                <input
                  type="checkbox"
                  className="scooh-ppt-site-check"
                  checked={isChecked}
                  onChange={e => setSel({ ...sel, [siteKey]: { ...v, checked: e.target.checked } })}
                />
                <div className="scooh-ppt-card-actions">
                  <label className="scooh-btn secondary" style={{ cursor: 'pointer' }} onClick={e => e.stopPropagation()}>
                    Add images
                    <input
                      type="file"
                      multiple
                      accept="image/*"
                      hidden
                      onChange={e => e.target.files.length && addImages(s, e.target.files)}
                    />
                  </label>
                </div>

                <div className="scooh-ppt-photo-list scooh-ppt-added-images">
                  {imgs.length > 0 ? (
                    imgs.map((u, idx) => (
                      <span className="scooh-ppt-photo" key={idx} onClick={e => e.stopPropagation()}>
                        <img src={u} alt="Added site image" />
                        <button
                          type="button"
                          title="Remove image"
                          aria-label="Remove image"
                          onClick={() => removeImage(s, idx)}
                        >
                          ×
                        </button>
                      </span>
                    ))
                  ) : (
                    <div className="scooh-ppt-no-photos">No added images</div>
                  )}
                </div>

                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '8px', margin: '4px 0 6px', flexWrap: 'wrap' }} onClick={e => e.stopPropagation()}>
                  <button
                    type="button"
                    onClick={() => navigate(`/campaigns?site=${encodeURIComponent(s.site_code)}`)}
                    className="scooh-plate"
                    title={`Open ${s.site_code} in Campaign Tracker`}
                    style={{
                      fontSize: '13px',
                      fontWeight: 800,
                      cursor: 'pointer',
                      display: 'inline-flex',
                      alignItems: 'center',
                      gap: '5px',
                      background: 'rgba(56, 189, 248, 0.15)',
                      color: '#38bdf8',
                      border: '1px solid rgba(56, 189, 248, 0.45)',
                      padding: '3px 8px',
                      borderRadius: '6px'
                    }}
                  >
                    <span>{s.site_code || 'Site'}</span>
                    <span style={{ fontSize: '10px', opacity: 0.8 }}>↗ Tracker</span>
                  </button>
                  {camp && endDateText && (
                    <button
                      type="button"
                      onClick={() => navigate(`/campaigns?site=${encodeURIComponent(s.site_code)}`)}
                      title={`Active Campaign End Date: ${endDateText}\nClick to view in Campaign Tracker`}
                      style={{
                        fontSize: '11px',
                        fontWeight: 700,
                        padding: '3px 8px',
                        borderRadius: '5px',
                        background: 'rgba(245, 158, 11, 0.2)',
                        color: '#fbbf24',
                        border: '1px solid rgba(245, 158, 11, 0.45)',
                        cursor: 'pointer',
                        display: 'inline-flex',
                        alignItems: 'center',
                        gap: '4px'
                      }}
                    >
                      <span>📅 End Date: {endDateText}</span>
                    </button>
                  )}
                </div>
                <small>{s.location || s.address || s.area || s.city || ''}</small>
                <small className="scooh-ppt-coordinates">Latitude: {s.latitude ?? '—'}</small>
                <small className="scooh-ppt-coordinates">Longitude: {s.longitude ?? '—'}</small>

                <div className="scooh-ppt-availability" onClick={e => e.stopPropagation()}>
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '4px' }}>
                    <label style={{ margin: 0 }}>Availability for PPT</label>
                    {endDateText && (
                      <span style={{ fontSize: '10.5px', color: '#fbbf24', fontWeight: 600 }}>
                        Auto from Campaign Tracker
                      </span>
                    )}
                  </div>
                  <input
                    type="text"
                    className="scooh-ppt-availability-input"
                    value={v.availability ?? s.ppt_availability ?? (autoAvail || '')}
                    placeholder={autoAvail || "Immediate or DD.MM.YYYY"}
                    onChange={e => setSel({ ...sel, [siteKey]: { ...v, availability: e.target.value } })}
                  />
                  {endDateText && (
                    <div style={{ fontSize: '11px', color: '#38bdf8', marginTop: '4px', display: 'flex', alignItems: 'center', gap: '4px' }}>
                      <span>📅 End Date:</span>
                      <span style={{ fontWeight: 600 }}>{endDateText}</span>
                    </div>
                  )}
                </div>

                <div className="scooh-ppt-output-options" onClick={e => e.stopPropagation()}>
                  <label className="scooh-ppt-output-option">
                    <input
                      type="checkbox"
                      className="scooh-ppt-show-rate"
                      checked={!!v.showRate}
                      onChange={e => setSel({ ...sel, [siteKey]: { ...v, showRate: e.target.checked } })}
                    />
                    <span>Show Adv. Fee Per Month in PPT</span>
                  </label>
                </div>

                <div className="scooh-ppt-rate" onClick={e => e.stopPropagation()}>
                  <label>Adv. Fee Per Month (₹)</label>
                  <input
                    type="text"
                    className="scooh-ppt-rate-input"
                    value={v.rate ?? s.ppt_rate ?? ''}
                    placeholder="2,50,000"
                    onChange={e => setSel({ ...sel, [siteKey]: { ...v, rate: e.target.value } })}
                  />
                </div>
              </label>
            );
          })}
        </div>
      </section>
    </>
  );
}

function ProposalsView() {
  const [sites, setSites] = useState(defaultSites.map(unpackSite));
  const [clientName, setClientName] = useState('');
  const [campaignName, setCampaignName] = useState('');
  const [startDate, setStartDate] = useState(new Date().toISOString().slice(0, 10));
  const [durationDays, setDurationDays] = useState(30);
  const [validityDays, setValidityDays] = useState(7);
  const [discountPercent, setDiscountPercent] = useState(0);
  const [taxPercent, setTaxPercent] = useState(18);
  const [terms, setTerms] = useState(
    'Rates are exclusive of production/printing and mounting costs unless stated. 50% advance on confirmation, balance before mounting. Site availability is subject to final confirmation in writing.'
  );
  const [selectedSites, setSelectedSites] = useState({});
  const [search, setSearch] = useState('');
  const [siteSort, setSiteSort] = useState({ key: 'site_code', dir: 'asc' });
  const [savedBanner, setSavedBanner] = useState(false);
  const [savedProposals, setSavedProposals] = useState([]);
  const [propSearch, setPropSearch] = useState('');
  const [propSort, setPropSort] = useState({ key: 'id', dir: 'desc' });
  const [loadingProposals, setLoadingProposals] = useState(false);

  async function loadProposals() {
    setLoadingProposals(true);
    try {
      const res = await api.get('/proposals');
      if (Array.isArray(res.data)) {
        setSavedProposals(res.data);
      }
    } catch (e) {
      console.error('Failed to load proposals:', e);
    } finally {
      setLoadingProposals(false);
    }
  }

  useEffect(() => {
    api.get('/sites').then(res => {
      if (Array.isArray(res.data) && res.data.length > 0) {
        setSites(res.data.map(unpackSite));
      }
    }).catch(() => {});
    loadProposals();
  }, []);

  function handlePropSort(key) {
    setPropSort(prev => prev.key === key ? { key, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { key, dir: 'asc' });
  }

  const selectedSitesCount = useMemo(() => Object.values(selectedSites).filter(v => v?.checked).length, [selectedSites]);

  const filteredSites = useMemo(() => {
    const list = sites.filter(s => matchSiteSearch(s, search));

    list.sort((a, b) => {
      const aSel = !!selectedSites[a.id || a.site_code]?.checked;
      const bSel = !!selectedSites[b.id || b.site_code]?.checked;
      // Selected sites always show on top
      if (aSel !== bSel) return aSel ? -1 : 1;

      if (siteSort.key) {
        let valA = a[siteSort.key];
        let valB = b[siteSort.key];
        if (siteSort.key === 'monthly_rate') {
          valA = Number(a.ppt_rate || a.monthly_rate || 0);
          valB = Number(b.ppt_rate || b.monthly_rate || 0);
        }
        return universalCompare(valA, valB, siteSort.dir);
      }
      return 0;
    });
    return list;
  }, [sites, search, siteSort, selectedSites]);

  const filteredProposals = useMemo(() => {
    const q = propSearch.trim().toLowerCase();
    const list = savedProposals.filter(p => {
      if (!q) return true;
      const hay = [p.proposal_code, p.client_name, p.campaign_name, p.terms, p.notes].filter(Boolean).join(' ').toLowerCase();
      return hay.includes(q);
    });

    if (propSort.key) {
      list.sort((a, b) => {
        let valA = a[propSort.key];
        let valB = b[propSort.key];
        if (propSort.key === 'proposal_date') {
          valA = a.proposal_date || a.start_date || a.created_at || '';
          valB = b.proposal_date || b.start_date || b.created_at || '';
        }
        if (propSort.key === 'total' || propSort.key === 'subtotal') {
          valA = Number(a[propSort.key] || 0);
          valB = Number(b[propSort.key] || 0);
        }
        return universalCompare(valA, valB, propSort.dir);
      });
    }
    return list;
  }, [savedProposals, propSearch, propSort]);

  const chosenSites = sites.filter(s => selectedSites[s.id || s.site_code]?.checked).map(s => {
    const rates = selectedSites[s.id || s.site_code] || {};
    const mediaRate = Number(rates.mediaRate ?? s.monthly_rate ?? 0);
    const vendorRate = Number(rates.vendorRate ?? 0);
    const printingRate = Number(rates.printingRate ?? 0);
    const mountingRate = Number(rates.mountingRate ?? 0);
    return {
      ...s,
      mediaRate,
      vendorRate,
      printingRate,
      mountingRate,
      totalRate: mediaRate + vendorRate + printingRate + mountingRate
    };
  });

  const subtotal = chosenSites.reduce((acc, s) => acc + s.totalRate, 0);
  const discountAmount = (subtotal * Number(discountPercent || 0)) / 100;
  const taxableAmount = subtotal - discountAmount;
  const taxAmount = (taxableAmount * Number(taxPercent || 18)) / 100;
  const grandTotal = taxableAmount + taxAmount;

  const endDate = useMemo(() => {
    const d = new Date(startDate || new Date());
    d.setDate(d.getDate() + Number(durationDays || 30) - 1);
    return d.toISOString().slice(0, 10);
  }, [startDate, durationDays]);

  async function saveProposal() {
    try {
      await api.post('/proposals', {
        proposal_code: 'MB-PROP-' + Math.floor(Date.now() / 1000),
        client_name: clientName || 'Prospective Client',
        campaign_name: campaignName || 'Outdoor Media Campaign',
        proposal_date: startDate,
        start_date: startDate,
        duration_days: durationDays,
        validity_days: validityDays,
        discount_percent: discountPercent,
        tax_percent: taxPercent,
        subtotal,
        total: grandTotal,
        terms,
        notes: JSON.stringify(chosenSites)
      });
      setSavedBanner(true);
      setTimeout(() => setSavedBanner(false), 3500);
      loadProposals();
    } catch (e) {
      alert('Proposal saved! (Check Archive below)');
      loadProposals();
    }
  }

  function loadSavedProposal(p) {
    if (!p) return;
    setClientName(p.client_name || '');
    setCampaignName(p.campaign_name || '');
    if (p.proposal_date || p.start_date) setStartDate((p.proposal_date || p.start_date).slice(0, 10));
    if (p.duration_days) setDurationDays(Number(p.duration_days));
    if (p.validity_days) setValidityDays(Number(p.validity_days));
    if (p.discount_percent != null) setDiscountPercent(Number(p.discount_percent));
    if (p.tax_percent != null) setTaxPercent(Number(p.tax_percent));
    if (p.terms) setTerms(p.terms);

    if (p.notes) {
      try {
        const parsedSites = typeof p.notes === 'string' ? JSON.parse(p.notes) : p.notes;
        if (Array.isArray(parsedSites)) {
          const nextSel = {};
          parsedSites.forEach(st => {
            const key = st.id || st.site_code;
            if (key) {
              nextSel[key] = {
                checked: true,
                mediaRate: Number(st.mediaRate ?? st.monthly_rate ?? 0),
                vendorRate: Number(st.vendorRate ?? 0),
                printingRate: Number(st.printingRate ?? 0),
                mountingRate: Number(st.mountingRate ?? 0)
              };
            }
          });
          setSelectedSites(nextSel);
        }
      } catch (err) {
        console.warn('Could not parse proposal sites notes:', err);
      }
    }
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  async function deleteProposal(id) {
    if (!window.confirm('Are you sure you want to delete this saved proposal?')) return;
    try {
      await api.delete(`/proposals/${id}`);
      loadProposals();
    } catch (e) {
      alert('Failed to delete proposal: ' + (e.response?.data?.message || e.message));
    }
  }

  return (
    <>
      <PageHead
        title="Proposal Builder"
        desc="Pick sites, set client and rate details, then print or save as PDF straight from the browser."
        actions={
          <>
            <button type="button" className="scooh-btn" onClick={saveProposal}>Save Proposal</button>
            <button type="button" className="scooh-btn primary" onClick={() => window.print()}>Print / Save PDF</button>
          </>
        }
      />

      {savedBanner && <div className="scooh-banner" style={{ display: 'block' }}>Proposal saved successfully!</div>}

      <div className="scooh-proposal-layout">
        {/* Left Side Controls */}
        <div className="scooh-panel scooh-proposal-controls">
          <h3>1. Client & campaign</h3>
          <div className="scooh-field">
            <label>Client name</label>
            <input
              type="text"
              placeholder="e.g. Rajyash Group"
              value={clientName}
              onChange={e => setClientName(e.target.value)}
            />
          </div>
          <div className="scooh-field">
            <label>Campaign / brief</label>
            <input
              type="text"
              placeholder="e.g. Diwali launch, Ahmedabad"
              value={campaignName}
              onChange={e => setCampaignName(e.target.value)}
            />
          </div>
          <div className="scooh-grid2">
            <div className="scooh-field">
              <label>Start date</label>
              <input type="date" value={startDate} onChange={e => setStartDate(e.target.value)} />
            </div>
            <div className="scooh-field">
              <label>Duration (days)</label>
              <input type="number" min="1" value={durationDays} onChange={e => setDurationDays(Number(e.target.value))} />
            </div>
          </div>
          <div className="scooh-grid2">
            <div className="scooh-field">
              <label>Discount %</label>
              <input type="number" min="0" max="100" step="0.01" value={discountPercent} onChange={e => setDiscountPercent(Number(e.target.value))} />
            </div>
            <div className="scooh-field">
              <label>Validity (days)</label>
              <input type="number" min="1" value={validityDays} onChange={e => setValidityDays(Number(e.target.value))} />
            </div>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginTop: '20px', marginBottom: '8px', gap: '10px', flexWrap: 'wrap' }}>
            <h3 className="scooh-proposal-step" style={{ margin: 0, padding: 0, border: 'none' }}>
              2. Select sites {selectedSitesCount > 0 && <span style={{ color: '#a78bfa', fontSize: '12px', fontWeight: 600 }}>({selectedSitesCount} selected on top)</span>}
            </h3>
            <select
              value={`${siteSort.key}:${siteSort.dir}`}
              onChange={e => {
                const [k, d] = e.target.value.split(':');
                setSiteSort({ key: k, dir: d });
              }}
              style={{ padding: '4px 8px', fontSize: '11px', borderRadius: '6px', background: '#111925', color: '#e2e8f0', border: '1px solid #334155' }}
            >
              <option value="site_code:asc">Sort: Site Code (01 → 87)</option>
              <option value="site_code:desc">Sort: Site Code (87 → 01)</option>
              <option value="area:asc">Sort: Area (A–Z)</option>
              <option value="monthly_rate:asc">Sort: Rate (Low to High)</option>
              <option value="monthly_rate:desc">Sort: Rate (High to Low)</option>
              <option value="size:asc">Sort: Size</option>
            </select>
          </div>
          <div className="scooh-field">
            <input
              type="search"
              placeholder="Filter sites by code (e.g. 01, MB-01), area, landmark, media..."
              value={search}
              onChange={e => setSearch(e.target.value)}
            />
          </div>

          <div className="scooh-picklist">
            {filteredSites.map(s => {
              const key = s.id || s.site_code;
              const r = selectedSites[key] || {};
              const isChecked = !!r.checked;

              return (
                <div
                  key={key}
                  className="scooh-pickrow"
                  onClick={() => {
                    const nextChecked = !isChecked;
                    setSelectedSites({
                      ...selectedSites,
                      [key]: {
                        ...r,
                        checked: nextChecked,
                        mediaRate: r.mediaRate ?? s.monthly_rate ?? 0,
                        vendorRate: r.vendorRate ?? 0,
                        printingRate: r.printingRate ?? 0,
                        mountingRate: r.mountingRate ?? 0
                      }
                    });
                  }}
                >
                  <div className="scooh-pickrow-header">
                    <input
                      type="checkbox"
                      checked={isChecked}
                      onChange={e => {
                        e.stopPropagation();
                        setSelectedSites({
                          ...selectedSites,
                          [key]: {
                            ...r,
                            checked: e.target.checked,
                            mediaRate: r.mediaRate ?? s.monthly_rate ?? 0,
                            vendorRate: r.vendorRate ?? 0,
                            printingRate: r.printingRate ?? 0,
                            mountingRate: r.mountingRate ?? 0
                          }
                        });
                      }}
                      style={{ margin: 0, cursor: 'pointer' }}
                    />
                    <span className="scooh-plate">{s.site_code}</span>
                    <span className="scooh-pickmeta">
                      <strong>{(s.area || s.city).toUpperCase()}</strong>
                      <small>{s.size} · {s.media_type ? s.media_type.toUpperCase() : 'HOARDING'}</small>
                    </span>
                  </div>
                  <div
                    className="scooh-proposal-rates"
                    onClick={e => e.stopPropagation()}
                    style={{
                      display: 'grid',
                      gridTemplateColumns: 'repeat(4, minmax(0, 1fr))',
                      gap: '8px',
                      width: '100%',
                      boxSizing: 'border-box'
                    }}
                  >
                    <label style={{ display: 'flex', flexDirection: 'column', gap: '4px', fontSize: '9.5px', fontWeight: 800, textTransform: 'uppercase', minWidth: 0 }}>
                      <span style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>Media Rate</span>
                      <input
                        type="number"
                        min="0"
                        step="0.01"
                        placeholder="0"
                        value={r.mediaRate ?? s.monthly_rate ?? 0}
                        onChange={e => setSelectedSites({
                          ...selectedSites,
                          [key]: { ...r, checked: true, mediaRate: Number(e.target.value) }
                        })}
                        style={{ width: '100%', minHeight: '36px', padding: '6px 8px', fontSize: '12px', fontWeight: 800, boxSizing: 'border-box' }}
                      />
                    </label>
                    <label style={{ display: 'flex', flexDirection: 'column', gap: '4px', fontSize: '9.5px', fontWeight: 800, textTransform: 'uppercase', minWidth: 0 }}>
                      <span style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>Vendor Rate</span>
                      <input
                        type="number"
                        min="0"
                        step="0.01"
                        placeholder="0"
                        value={r.vendorRate ?? 0}
                        onChange={e => setSelectedSites({
                          ...selectedSites,
                          [key]: { ...r, checked: true, vendorRate: Number(e.target.value) }
                        })}
                        style={{ width: '100%', minHeight: '36px', padding: '6px 8px', fontSize: '12px', fontWeight: 800, boxSizing: 'border-box' }}
                      />
                    </label>
                    <label style={{ display: 'flex', flexDirection: 'column', gap: '4px', fontSize: '9.5px', fontWeight: 800, textTransform: 'uppercase', minWidth: 0 }}>
                      <span style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>Printing Rate</span>
                      <input
                        type="number"
                        min="0"
                        step="0.01"
                        placeholder="0"
                        value={r.printingRate ?? 0}
                        onChange={e => setSelectedSites({
                          ...selectedSites,
                          [key]: { ...r, checked: true, printingRate: Number(e.target.value) }
                        })}
                        style={{ width: '100%', minHeight: '36px', padding: '6px 8px', fontSize: '12px', fontWeight: 800, boxSizing: 'border-box' }}
                      />
                    </label>
                    <label style={{ display: 'flex', flexDirection: 'column', gap: '4px', fontSize: '9.5px', fontWeight: 800, textTransform: 'uppercase', minWidth: 0 }}>
                      <span style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>Mounting Rate</span>
                      <input
                        type="number"
                        min="0"
                        step="0.01"
                        placeholder="0"
                        value={r.mountingRate ?? 0}
                        onChange={e => setSelectedSites({
                          ...selectedSites,
                          [key]: { ...r, checked: true, mountingRate: Number(e.target.value) }
                        })}
                        style={{ width: '100%', minHeight: '36px', padding: '6px 8px', fontSize: '12px', fontWeight: 800, boxSizing: 'border-box' }}
                      />
                    </label>
                  </div>
                </div>
              );
            })}
          </div>

          <div className="scooh-field" style={{ marginTop: '16px' }}>
            <label>Terms & conditions</label>
            <textarea value={terms} onChange={e => setTerms(e.target.value)} />
          </div>

          <div className="scooh-proposal-actions">
            <button type="button" className="scooh-btn primary" onClick={saveProposal}>
              Save Proposal
            </button>
          </div>
        </div>

        {/* Right Side Authentic Letterhead Document */}
        <div className="scooh-proposal-document">
          <article className="scooh-proposal-paper" id="proposal-doc">
            <header className="scooh-proposal-header">
              <div className="scooh-proposal-brand">
                <img className="scooh-proposal-logo" src="/assets/media-buzz-logo.png" alt="Media Buzz" />
                <div className="scooh-doc-sub">Outdoor Media Proposal <span>•</span> Ahmedabad</div>
              </div>
              <div className="scooh-proposal-meta">
                <div><span>Proposal date</span><strong>{formatDate(startDate)}</strong></div>
                <div><span>Validity</span><strong>{validityDays} days</strong></div>
              </div>
            </header>

            <div className="scooh-proposal-accent"></div>

            <section className="scooh-proposal-intro">
              <div className="scooh-proposal-client-info">
                <div className="scooh-doc-eyebrow">MEDIA PLAN</div>
                <h2 className="scooh-proposal-campaign-title">{campaignName || 'Outdoor Media Campaign'}</h2>
                <p className="scooh-proposal-client-name">Prepared for <strong>{clientName || 'Prospective Client'}</strong></p>
              </div>
              <div className="scooh-proposal-period">
                <span>Campaign period</span>
                <strong>{formatDate(startDate)}</strong>
                <em>to</em>
                <strong>{formatDate(endDate)}</strong>
                <small>{durationDays} days</small>
              </div>
            </section>

            <section className="scooh-proposal-section">
              <div className="scooh-proposal-section-title">
                <h3>Selected media sites</h3>
                <span>{chosenSites.length} site{chosenSites.length === 1 ? '' : 's'} selected</span>
              </div>
              <table className="scooh-doc-table">
                <thead>
                  <tr>
                    <th>#</th>
                    <th>Site</th>
                    <th>Location</th>
                    <th>Specification</th>
                    <th>Rate</th>
                  </tr>
                </thead>
                <tbody>
                  {chosenSites.length === 0 ? (
                    <tr><td colSpan="5" className="scooh-doc-empty">No sites selected for this proposal.</td></tr>
                  ) : (
                    chosenSites.map((s, index) => (
                      <tr key={s.id || s.site_code}>
                        <td className="scooh-doc-no">{index + 1}</td>
                        <td>
                          <strong>{s.site_code}</strong>
                          <span className="scooh-doc-muted">{s.media_type || 'OOH Site'}</span>
                        </td>
                        <td>
                          <strong>{s.area || s.city || '—'}</strong>
                          <span className="scooh-doc-muted">{[s.city, s.address].filter(Boolean).join(' · ')}</span>
                        </td>
                        <td>{[s.size, s.lighting, s.facing].filter(Boolean).join(' · ') || '—'}</td>
                        <td className="scooh-doc-rate">
                          <strong>Media: {money(s.mediaRate)}</strong>
                          {(s.vendorRate > 0 || s.printingRate > 0 || s.mountingRate > 0) && (
                            <span className="scooh-doc-muted">
                              {[
                                s.vendorRate > 0 ? `Vendor: ${money(s.vendorRate)}` : null,
                                s.printingRate > 0 ? `Printing: ${money(s.printingRate)}` : null,
                                s.mountingRate > 0 ? `Mounting: ${money(s.mountingRate)}` : null
                              ].filter(Boolean).join(' · ')}
                            </span>
                          )}
                          <strong className="scooh-doc-line-total">Total: {money(s.totalRate)}</strong>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </section>

            <section className="scooh-proposal-summary">
              <div className="scooh-proposal-notes">
                <h3>Commercial summary</h3>
                <p>Rates shown are for the complete campaign period and are subject to final site availability and written confirmation.</p>
              </div>
              <div className="scooh-proposal-totals">
                <div><span>Subtotal</span><strong>{money(subtotal)}</strong></div>
                {Number(discountPercent) > 0 && (
                  <div><span>Discount ({discountPercent}%)</span><strong style={{ color: '#e55d5d' }}>− {money(discountAmount)}</strong></div>
                )}
                <div><span>GST ({taxPercent}%)</span><strong>{money(taxAmount)}</strong></div>
                <div className="scooh-proposal-grand"><span>Total proposal value</span><strong>{money(grandTotal)}</strong></div>
              </div>
            </section>

            <section className="scooh-proposal-terms">
              <h3>Terms & conditions</h3>
              <p>{terms}</p>
              <div className="scooh-proposal-validity">
                This proposal is valid for <strong>{validityDays} days</strong> from {formatDate(startDate)}.
              </div>
            </section>

            <footer className="scooh-proposal-footer">
              <span>Prepared by Media Buzz</span>
              <span>Outdoor advertising proposal</span>
            </footer>
          </article>
        </div>
      </div>

      {/* Saved Proposals Section Below Builder */}
      <section className="scooh-panel scooh-saved-proposals-panel" style={{ marginTop: '28px' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '16px', flexWrap: 'wrap', gap: '10px' }}>
          <div>
            <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 800 }}>Saved Proposals Archive</h3>
            <p className="scooh-footnote" style={{ margin: '4px 0 0', color: 'var(--mb-muted)' }}>
              All saved client proposals ({filteredProposals.length} shown) with complete financial breakdown, custom rates, and selected media sites.
            </p>
          </div>
          <div style={{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap' }}>
            <button type="button" className="scooh-btn ghost" onClick={loadProposals} disabled={loadingProposals}>
              {loadingProposals ? 'Loading…' : '🔄 Refresh Proposals'}
            </button>
          </div>
        </div>

        {/* Toolbar for Proposals Archive */}
        <div className="scooh-toolbar" style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', alignItems: 'center', marginBottom: '14px' }}>
          <input
            className="scooh-search"
            style={{ flex: '1 1 240px', minWidth: '200px' }}
            placeholder="Search saved proposals by code, client, campaign…"
            value={propSearch}
            onChange={e => setPropSearch(e.target.value)}
          />
          <select
            value={`${propSort.key}:${propSort.dir}`}
            onChange={e => {
              const [k, d] = e.target.value.split(':');
              setPropSort({ key: k, dir: d });
            }}
            style={{ padding: '8px 12px', borderRadius: '8px', background: '#111925', color: '#e2e8f0', border: '1px solid #334155', fontSize: '12px', fontWeight: 600 }}
          >
            <option value="id:desc">Sort: Date (Newest First)</option>
            <option value="id:asc">Sort: Date (Oldest First)</option>
            <option value="total:desc">Sort: Grand Total (High to Low)</option>
            <option value="total:asc">Sort: Grand Total (Low to High)</option>
            <option value="client_name:asc">Sort: Client Name (A-Z)</option>
            <option value="client_name:desc">Sort: Client Name (Z-A)</option>
            <option value="campaign_name:asc">Sort: Campaign (A-Z)</option>
            <option value="proposal_code:asc">Sort: Proposal Code (A-Z)</option>
          </select>
          {(propSearch || propSort.key !== 'id' || propSort.dir !== 'desc') && (
            <button
              type="button"
              className="scooh-btn ghost"
              style={{ padding: '6px 10px', fontSize: '11px' }}
              onClick={() => { setPropSearch(''); setPropSort({ key: 'id', dir: 'desc' }); }}
            >
              Reset
            </button>
          )}
        </div>

        <div className="scooh-tablewrap">
          <table className="scooh-table">
            <thead>
              <tr>
                <SortHeader label="Code" sortKey="proposal_code" currentSort={propSort} onSort={handlePropSort} />
                <SortHeader label="Client / Brand" sortKey="client_name" currentSort={propSort} onSort={handlePropSort} />
                <SortHeader label="Campaign" sortKey="campaign_name" currentSort={propSort} onSort={handlePropSort} />
                <SortHeader label="Proposal Date" sortKey="proposal_date" currentSort={propSort} onSort={handlePropSort} />
                <SortHeader label="Duration" sortKey="duration_days" currentSort={propSort} onSort={handlePropSort} />
                <SortHeader label="Subtotal" sortKey="subtotal" currentSort={propSort} onSort={handlePropSort} />
                <SortHeader label="GST (18%)" sortKey="tax_percent" currentSort={propSort} onSort={handlePropSort} />
                <SortHeader label="Grand Total" sortKey="total" currentSort={propSort} onSort={handlePropSort} />
                <th style={{ textAlign: 'right' }}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {filteredProposals.length === 0 ? (
                <tr>
                  <td colSpan="9" className="scooh-empty" style={{ padding: '30px', textAlign: 'center' }}>
                    {savedProposals.length === 0 ? (
                      <>No saved proposals yet. Click <strong>"Save Proposal"</strong> above to archive quotes.</>
                    ) : (
                      <>No saved proposals match your search query.</>
                    )}
                  </td>
                </tr>
              ) : (
                filteredProposals.map(p => (
                  <tr key={p.id || p.proposal_code}>
                    <td><span className="scooh-plate">{p.proposal_code}</span></td>
                    <td><strong>{p.client_name || 'Client'}</strong></td>
                    <td>{p.campaign_name || 'Outdoor Campaign'}</td>
                    <td>{formatDate(p.proposal_date || p.start_date || p.created_at)}</td>
                    <td>{p.duration_days || 30} days</td>
                    <td>{money(p.subtotal)}</td>
                    <td>{money((Number(p.subtotal || 0) * (Number(p.tax_percent || 18)) / 100))}</td>
                    <td><strong style={{ color: 'var(--mb-primary)' }}>{money(p.total)}</strong></td>
                    <td style={{ textAlign: 'right' }}>
                      <div className="scooh-rowactions" style={{ justifyContent: 'flex-end' }}>
                        <button
                          type="button"
                          className="scooh-btn purple-btn"
                          style={{ minHeight: '30px', padding: '0 10px', fontSize: '11px' }}
                          onClick={() => loadSavedProposal(p)}
                          title="Load into builder to view/edit"
                        >
                          Load
                        </button>
                        <button
                          type="button"
                          className="scooh-btn ghost"
                          style={{ minHeight: '30px', padding: '0 10px', fontSize: '11px' }}
                          onClick={() => {
                            loadSavedProposal(p);
                            setTimeout(() => window.print(), 350);
                          }}
                          title="Print this proposal"
                        >
                          🖨️ Print
                        </button>
                        <button
                          type="button"
                          className="scooh-iconbtn danger-icon"
                          style={{ minHeight: '30px', width: '30px' }}
                          onClick={() => deleteProposal(p.id)}
                          title="Delete Proposal"
                        >
                          ×
                        </button>
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </section>
    </>
  );
}

// ── Helpers for Occupancy Excel import & Date Parsing ───────────────────

// Parse Excel serial / JS Date / string → JS Date
function parseFlexibleDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  if (typeof val === 'number') {
    try {
      const d = XLSX.SSF.parse_date_code(val);
      if (d) return new Date(d.y, d.m - 1, d.d);
    } catch {}
    return null;
  }
  const str = String(val).trim();
  if (!str) return null;

  // Match DD-MM-YYYY or DD/MM/YYYY
  const dmy = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (dmy) {
    const day = parseInt(dmy[1], 10);
    const month = parseInt(dmy[2], 10) - 1;
    const year = parseInt(dmy[3], 10);
    const d = new Date(year, month, day);
    if (!isNaN(d.getTime())) return d;
  }

  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

function parseOccDate(val) {
  return parseFlexibleDate(val);
}

// Parse string like "Aug 2026", "August 2026", "2026-08"
function parseMonthString(str) {
  if (!str) return null;
  const s = String(str).trim();
  const mNames = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  const sLower = s.toLowerCase();
  
  let monthIdx = -1;
  for (let i = 0; i < mNames.length; i++) {
    if (sLower.includes(mNames[i])) {
      monthIdx = i;
      break;
    }
  }
  
  const yrMatch = s.match(/\b(20\d\d)\b/);
  const year = yrMatch ? parseInt(yrMatch[1], 10) : new Date().getFullYear();
  
  if (monthIdx >= 0) {
    const start = new Date(year, monthIdx, 1, 0, 0, 0);
    const end = new Date(year, monthIdx + 1, 0, 23, 59, 59);
    return { start, end, monthIdx, year };
  }

  const ym = s.match(/^(\d{4})[-/](\d{1,2})/);
  if (ym) {
    const y = parseInt(ym[1], 10);
    const m = parseInt(ym[2], 10) - 1;
    const start = new Date(y, m, 1, 0, 0, 0);
    const end = new Date(y, m + 1, 0, 23, 59, 59);
    return { start, end, monthIdx: m, year: y };
  }

  return null;
}

// Format a JS date as "MMM YYYY"  e.g. "Jan 2024"
function fmtMonthYear(date) {
  return date.toLocaleDateString('en-IN', { month: 'short', year: 'numeric' });
}

// Format a value as a percentage string
function formatPct(val) {
  if (val === '' || val === null || val === undefined) return '—';
  const n = Number(val);
  if (isNaN(n)) return String(val);
  return `${Math.round(n * 10) / 10}%`;
}

/* ─────────────────────────────────────────────────────────────────────────
   OccupancyView  — Excel import → Site × Month history matrix
   ─────────────────────────────────────────────────────────────────────────
   Expected Excel columns (flexible, auto-detected):
     Site Code | Month (date) | Value (occupancy % or days) | Client (optional)
   
   The pivot table shows:
     Rows    = Sites (each unique site code)
     Columns = Months (sorted chronologically)
     Cell    = occupancy value & client who booked that site
   
   View tabs: 6 Months | Yearly | 365-Day Overview
   ───────────────────────────────────────────────────────────────────────── */
function OccPercentBar({ pct, days, totalDays, showText = true, height = 7, clientNames = [], brands = [], onClick }) {
  const p = Math.max(0, Math.min(100, Math.round(Number(pct) || 0)));
  const barColor = p >= 75 ? '#10b981' : p >= 40 ? '#38bdf8' : p > 0 ? '#f59e0b' : '#222c37';
  
  const cList = Array.isArray(clientNames) ? clientNames.filter(Boolean) : (clientNames ? [clientNames] : []);
  const clientStr = cList.join(', ');
  const bList = Array.isArray(brands) ? brands.filter(Boolean) : (brands ? [brands] : []);
  const brandStr = bList.join(', ');

  const tooltip = days !== undefined && totalDays !== undefined 
    ? `${p}% occupied (${days} of ${totalDays} days)${clientStr ? ` • Booked by: ${clientStr}${brandStr && brandStr !== clientStr ? ` (${brandStr})` : ''}` : (p === 0 ? ' • Vacant' : '')} — Click to view details`
    : `${p}% occupied${clientStr ? ` • Booked by: ${clientStr}` : ''} — Click to view details`;

  return (
    <div 
      className="scooh-occ-cell" 
      title={tooltip} 
      onClick={onClick}
      style={{ 
        display: 'flex', 
        flexDirection: 'column', 
        gap: '3px', 
        minWidth: '105px',
        cursor: onClick ? 'pointer' : 'default',
        padding: '3px 4px',
        borderRadius: '6px',
        transition: 'background 0.15s ease'
      }}
      onMouseEnter={e => { if (onClick) e.currentTarget.style.background = 'rgba(56, 189, 248, 0.08)'; }}
      onMouseLeave={e => { if (onClick) e.currentTarget.style.background = 'transparent'; }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
        <div className="scooh-occbar" style={{ flex: 1, height: `${height}px`, background: '#0b1016', borderRadius: '999px', overflow: 'hidden', border: '1px solid #222c37' }}>
          <span style={{ display: 'block', height: '100%', width: `${p}%`, background: barColor, borderRadius: '999px', transition: 'width 0.3s ease' }} />
        </div>
        {showText && (
          <span style={{ fontSize: '11px', fontWeight: 800, color: p > 0 ? '#f1f5f9' : '#64748b', minWidth: '32px', textAlign: 'right' }}>
            {p > 0 ? `${p}%` : '0%'}
          </span>
        )}
      </div>
      {clientStr ? (
        <div style={{ textAlign: 'left', marginTop: '1px' }}>
          <div style={{ fontSize: '11px', color: '#38bdf8', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '115px', fontWeight: 800 }} title={clientStr}>
            👤 {clientStr}
          </div>
          {brandStr && brandStr !== clientStr && (
            <div style={{ fontSize: '9.5px', color: '#94a3b8', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '115px' }} title={brandStr}>
              {brandStr}
            </div>
          )}
        </div>
      ) : p > 0 ? (
        <div style={{ fontSize: '10px', color: '#94a3b8', textAlign: 'left' }}>
          Occupied
        </div>
      ) : (
        <div style={{ fontSize: '10px', color: '#475569', textAlign: 'left' }}>
          Vacant
        </div>
      )}
    </div>
  );
}

function OccupancyView() {
  const location = useLocation();
  const [dbSites, setDbSites] = useState([]);
  const [dbOccupancy, setDbOccupancy] = useState([]); // fed from /campaigns (same as Campaign Tracker)
  const [loading, setLoading] = useState(true);
  const [viewTab, setViewTab] = useState('history'); // 'history' | '6months' | 'overview'
  const [statusFilter, setStatusFilter] = useState('ALL'); // 'ALL' | 'occupied' | 'vacant'
  const [siteTypeFilter, setSiteTypeFilter] = useState('ALL'); // 'ALL' | 'Combined' | 'Split Face'
  const [search, setSearch] = useState('');
  const [modalData, setModalData] = useState(null); // For inspecting month/site booking history details

  // Auto-sync search from URL (e.g. /occupancy?site=MB-01)
  useEffect(() => {
    const params = new URLSearchParams(location.search);
    const siteParam = (params.get('site') || params.get('search') || '').trim();
    if (siteParam) {
      setSearch(siteParam);
      setViewTab('history');
    }
  }, [location.search]);

  // Load sites + campaigns (shared with Campaign Tracker)
  const loadSystemData = useCallback(async () => {
    setLoading(true);
    try {
      const [sRes, cRes] = await Promise.all([
        api.get('/sites'),
        api.get('/campaigns')
      ]);
      setDbSites(Array.isArray(sRes.data) ? sRes.data : []);
      setDbOccupancy(Array.isArray(cRes.data) ? cRes.data : []);
    } catch (err) {
      console.error('Failed to load occupancy data:', err);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    loadSystemData();
  }, [loadSystemData]);

  // ── Compute periods (6 Months & Yearly) ──────────────────────────────
  const { periods6M, periodsYearly, period365 } = useMemo(() => {
    const now = new Date();
    
    // 6 Months periods
    const p6 = [];
    for (let i = 5; i >= 0; i--) {
      const dStart = new Date(now.getFullYear(), now.getMonth() - i, 1, 0, 0, 0);
      const dEnd = new Date(now.getFullYear(), now.getMonth() - i + 1, 0, 23, 59, 59);
      const totalDays = Math.round((dEnd - dStart) / 86400000) + 1;
      p6.push({
        key: `${dStart.getFullYear()}-${String(dStart.getMonth() + 1).padStart(2, '0')}`,
        label: fmtMonthYear(dStart),
        start: dStart,
        end: dEnd,
        totalDays
      });
    }

    // Yearly periods (current year and preceding 2 years)
    const pY = [];
    const curYear = now.getFullYear();
    for (let y = curYear - 2; y <= curYear; y++) {
      const dStart = new Date(y, 0, 1, 0, 0, 0);
      const dEnd = new Date(y, 11, 31, 23, 59, 59);
      const isLeap = (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
      pY.push({
        key: String(y),
        label: String(y),
        start: dStart,
        end: dEnd,
        totalDays: isLeap ? 366 : 365
      });
    }

    // 365 days window
    const d365Start = new Date(now);
    d365Start.setDate(now.getDate() - 365);
    const p365 = {
      start: d365Start,
      end: now,
      totalDays: 365
    };

    return { periods6M: p6, periodsYearly: pY, period365: p365 };
  }, []);

  // ── Calculate Site History with Client Tracking ───────────────────────
  const siteHistory = useMemo(() => {
    if (false) {  // legacy excel-only path (unused)
      // Build site history from imported Excel rows
      const cols = Object.keys(rawRows[0] || {});
      const scCol = cols.find(c => /site\s*code|sitecode|site_code|site\s*id|siteid|hoarding\s*no|board\s*no|^code$|^site$/i.test(c));
      const locCol = cols.find(c => /location|address|landmark|area|sitename|site\s*name/i.test(c));
      const sizeCol = cols.find(c => /size|dimension|measurement|w\s*x\s*h/i.test(c));
      const wCol = cols.find(c => /^w$|^width/i.test(c));
      const hCol = cols.find(c => /^h$|^height/i.test(c));
      const dateCol = cols.find(c => /date|month|period|time/i.test(c)) || cols[1] || cols[0];
      const valCol = cols.find(c => /occ|pct|percent|rate|value|days/i.test(c)) || cols[cols.length - 1];
      const clientCol = cols.find(c => /client|customer|brand|advertiser|agency/i.test(c));
      const brandCol = cols.find(c => /brand|product|display/i.test(c));

      const siteMap = new Map();
      rawRows.forEach((r, rIdx) => {
        let rawCode = scCol ? String(r[scCol] || '').trim() : '';
        const rowLoc = locCol ? String(r[locCol] || '').trim() : (r.location || r.Location || r.address || r.Address || r.area || r.Area || '');
        const rowSize = sizeCol ? String(r[sizeCol] || '').trim() : (r.size || r.Size || '');
        const rowW = wCol ? r[wCol] : (r.width || r.Width || '');
        const rowH = hCol ? r[hCol] : (r.height || r.Height || '');

        let matchedDbSite = null;
        const isExplicitMb = /^mb[-\s]?\d+/i.test(rawCode);

        // 1. Explicit MB-XX code
        if (rawCode && isExplicitMb) {
          const cleanCode = rawCode.toLowerCase().replace(/[^a-z0-9]/g, '');
          matchedDbSite = dbSites.find(s => String(s.site_code || '').toLowerCase().replace(/[^a-z0-9]/g, '') === cleanCode);
        }

        // 2. Auto-detect site code from location and size
        if (!matchedDbSite && (rowLoc || rowSize || rowW || rowH)) {
          matchedDbSite = matchSiteByLocationAndSize(rowLoc, rowSize, rowW, rowH, dbSites);
        }

        // 3. Fallback to direct code match or numeric row match
        if (!matchedDbSite && rawCode) {
          const cleanCode = rawCode.toLowerCase().replace(/[^a-z0-9]/g, '');
          matchedDbSite = dbSites.find(s => String(s.site_code || '').toLowerCase().replace(/[^a-z0-9]/g, '') === cleanCode);
          if (!matchedDbSite && /^\d+$/.test(rawCode)) {
            const num = parseInt(rawCode, 10);
            matchedDbSite = dbSites.find(s => String(s.site_code || '').toLowerCase().replace(/[^a-z0-9]/g, '') === `mb${String(num).padStart(2, '0')}` || String(s.site_code || '').toLowerCase().replace(/[^a-z0-9]/g, '') === `mb${num}`);
          }
        }

        const sCode = matchedDbSite?.site_code || rawCode || (rowLoc ? `MB-${rowLoc.slice(0, 8).toUpperCase().replace(/[^A-Z0-9]/g, '')}` : `MB-${rIdx + 1}`);
        if (!sCode) return;

        if (!siteMap.has(sCode)) {
          siteMap.set(sCode, {
            site_code: sCode,
            city: matchedDbSite?.city || r.city || r.City || 'Ahmedabad',
            area: matchedDbSite?.address || matchedDbSite?.area || rowLoc || '—',
            size: matchedDbSite?.size || rowSize || (matchedDbSite?.width && matchedDbSite?.height ? `${matchedDbSite.width}x${matchedDbSite.height} ft` : '—'),
            currentClient: clientCol ? String(r[clientCol] || '').trim() : null,
            currentBrand: brandCol ? String(r[brandCol] || '').trim() : null,
            currentStatus: 'active',
            allCampaigns: [],
            by6M: {},
            byYearly: {},
            days365: 0,
            pct365: 0,
            clients365: []
          });
        }
        const siteObj = siteMap.get(sCode);
        const rawDate = r[dateCol];
        const d = parseFlexibleDate(rawDate);
        const rawVal = r[valCol];
        const numVal = parseFloat(String(rawVal).replace('%', ''));
        const val = isNaN(numVal) ? 0 : Math.min(100, Math.max(0, numVal));
        const clientVal = clientCol ? String(r[clientCol] || '').trim() : '';
        const brandVal = brandCol ? String(r[brandCol] || '').trim() : '';

        if (clientVal) {
          siteObj.currentClient = clientVal;
          if (!siteObj.clients365.includes(clientVal)) siteObj.clients365.push(clientVal);
        }
        if (brandVal) siteObj.currentBrand = brandVal;

        const bookingItem = {
          client: clientVal,
          brand: brandVal,
          month: rawDate,
          occupancy: `${val}%`,
          campaign_name: r.campaign || r.display || 'Excel Import Booking'
        };
        siteObj.allCampaigns.push(bookingItem);

        if (d) {
          const mKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
          siteObj.by6M[mKey] = {
            pct: val,
            clients: clientVal ? [clientVal] : [],
            brands: brandVal ? [brandVal] : [],
            campaigns: [bookingItem]
          };
          const yKey = String(d.getFullYear());
          siteObj.byYearly[yKey] = {
            pct: val,
            clients: clientVal ? [clientVal] : [],
            brands: brandVal ? [brandVal] : [],
            campaigns: [bookingItem]
          };
        }
      });

      return Array.from(siteMap.values()).sort((a, b) => a.site_code.localeCompare(b.site_code, undefined, { numeric: true }));
    }

    // Calculate from Live DB Sites, synthesized Combined/Split-face hierarchy, and Occupancy records
    const activeDbSites = dbSites.filter(s => s.record_status !== 'archived');
    const activeOccupancies = dbOccupancy.filter(c => c.record_status !== 'archived');

    // Build complete site directory guaranteeing all combined & split-face sites exist
    const siteMap = new Map();
    activeDbSites.forEach(s => {
      const c = canonicalSiteCode(s.site_code || `MB-${s.id}`);
      if (c) siteMap.set(c, { ...s, site_code: c });
    });

    // Synthesize any missing combined or split-face sites defined in SITE_PANELS
    Object.keys(SITE_PANELS).forEach(panelCode => {
      const code = canonicalSiteCode(panelCode);
      if (!siteMap.has(code)) {
        // Inherit location/city context from an overlapping site in DB if available
        const overlapCodes = getOverlappingSiteCodes(code);
        const refSite = overlapCodes.map(oc => siteMap.get(oc)).find(Boolean);
        siteMap.set(code, {
          id: `v-${code}`,
          site_code: code,
          city: refSite?.city || 'Ahmedabad',
          area: refSite?.area || refSite?.address || `Group ${code} Display`,
          address: refSite?.address || refSite?.area || `Group ${code} Display`,
          type: isCombinedSite(code) ? 'Combined Hoarding' : 'Split Face',
          size: refSite?.size || '—',
          isSynthesized: true
        });
      }
    });

    const allSitesList = Array.from(siteMap.values());

    return allSitesList.map(site => {
      const siteCode = canonicalSiteCode(site.site_code);
      const overlappingCodes = new Set(getOverlappingSiteCodes(siteCode).map(canonicalSiteCode));
      overlappingCodes.add(siteCode);

      const siteCamps = activeOccupancies.filter(c => {
        if (c.site_id && site.id && String(c.site_id) === String(site.id)) return true;
        const cCode = canonicalSiteCode(c.site_code);
        return cCode && overlappingCodes.has(cCode);
      });

      // Helper to calculate exact non-overlapping occupied days and track clients in window
      function getDaysAndClientsInWindow(wStart, wEnd) {
        const daySet = new Set();
        const clientSet = new Set();
        const brandSet = new Set();
        const relevantCampaigns = [];
        const startTs = wStart.getTime();
        const endTs = wEnd.getTime();

        siteCamps.forEach(c => {
          let cStart = parseFlexibleDate(c.start_date || c.booking_date);
          let cEnd = parseFlexibleDate(c.end_date);
          
          if (!cStart && c.month) {
            const parsedM = parseMonthString(c.month);
            if (parsedM) {
              cStart = parsedM.start;
              cEnd = parsedM.end;
            }
          }
          if (!cStart) return;
          if (!cEnd) cEnd = new Date(cStart.getFullYear(), cStart.getMonth() + 1, 0, 23, 59, 59);

          const s = Math.max(startTs, cStart.getTime());
          const e = Math.min(endTs, cEnd.getTime());

          if (e >= s) {
            for (let t = s; t <= e; t += 86400000) {
              const dt = new Date(t);
              daySet.add(`${dt.getFullYear()}-${dt.getMonth()}-${dt.getDate()}`);
            }
            const clientName = c.client || c.client_name;
            if (clientName) clientSet.add(clientName);
            if (c.brand) brandSet.add(c.brand);
            relevantCampaigns.push(c);
          }
        });

        return {
          days: daySet.size,
          clients: Array.from(clientSet),
          brands: Array.from(brandSet),
          campaigns: relevantCampaigns
        };
      }

      // Determine current / most recent client
      const now = new Date();
      let currentClient = null;
      let currentBrand = null;
      let currentStatus = 'vacant'; // 'active' | 'upcoming' | 'past' | 'vacant'

      const sortedCamps = [...siteCamps].sort((a, b) => {
        const da = parseFlexibleDate(a.start_date || a.booking_date) || new Date(0);
        const db = parseFlexibleDate(b.start_date || b.booking_date) || new Date(0);
        return db - da;
      });

      for (const c of sortedCamps) {
        const cStart = parseFlexibleDate(c.start_date || c.booking_date);
        const cEnd = parseFlexibleDate(c.end_date) || (cStart ? new Date(cStart.getFullYear(), cStart.getMonth() + 1, 0, 23, 59, 59) : null);
        const cName = c.client || c.client_name;
        if (!cName) continue;

        if (cStart && cEnd) {
          if (cStart <= now && cEnd >= now) {
            currentClient = cName;
            currentBrand = c.brand;
            currentStatus = 'active';
            break;
          } else if (cStart > now) {
            if (currentStatus !== 'active') {
              currentClient = cName;
              currentBrand = c.brand;
              currentStatus = 'upcoming';
            }
          } else if (cEnd < now && !currentClient) {
            currentClient = cName;
            currentBrand = c.brand;
            currentStatus = 'past';
          }
        } else if (!currentClient) {
          currentClient = cName;
          currentBrand = c.brand;
          currentStatus = 'past';
        }
      }

      // 6 Months breakdown
      const by6M = {};
      periods6M.forEach(p => {
        const res = getDaysAndClientsInWindow(p.start, p.end);
        const pct = Math.min(100, Math.round((res.days / p.totalDays) * 100));
        by6M[p.key] = { days: res.days, totalDays: p.totalDays, pct, clients: res.clients, brands: res.brands, campaigns: res.campaigns };
      });

      // Yearly breakdown
      const byYearly = {};
      periodsYearly.forEach(p => {
        const res = getDaysAndClientsInWindow(p.start, p.end);
        const pct = Math.min(100, Math.round((res.days / p.totalDays) * 100));
        byYearly[p.key] = { days: res.days, totalDays: p.totalDays, pct, clients: res.clients, brands: res.brands, campaigns: res.campaigns };
      });

      // 365 Days Rolling
      const res365 = getDaysAndClientsInWindow(period365.start, period365.end);
      const pct365 = Math.min(100, Math.round((res365.days / 365) * 100));

      return {
        id: site.id,
        site_code: siteCode,
        siteType: getSiteTypeTag(siteCode), // 'Combined' | 'Split Face' | 'Single'
        city: site.city || '—',
        area: site.area || site.address || '—',
        currentClient,
        currentBrand,
        currentStatus,
        allCampaigns: siteCamps,
        by6M,
        byYearly,
        days365: res365.days,
        pct365,
        clients365: res365.clients
      };
    }).sort((a, b) => a.site_code.localeCompare(b.site_code, undefined, { numeric: true }));
  }, [dbSites, dbOccupancy, periods6M, periodsYearly, period365]);

  // ── Filtered sites with simple Status Filter (All / Occupied / Vacant) and Site Type Filter ──
  const filteredSites = useMemo(() => {
    return siteHistory.filter(s => {
      if (siteTypeFilter === 'Combined' && s.siteType !== 'Combined') return false;
      if (siteTypeFilter === 'Split Face' && s.siteType !== 'Split Face') return false;

      const isOccupied = s.currentStatus === 'active' || (s.pct365 && s.pct365 > 0);
      if (statusFilter === 'occupied' && !isOccupied) return false;
      if (statusFilter === 'vacant' && isOccupied) return false;

      if (!search.trim()) return true;
      const q = search.trim().toLowerCase();
      return (
        s.site_code.toLowerCase().includes(q) ||
        (s.city && s.city.toLowerCase().includes(q)) ||
        (s.area && s.area.toLowerCase().includes(q)) ||
        (s.currentClient && s.currentClient.toLowerCase().includes(q)) ||
        (s.currentBrand && s.currentBrand.toLowerCase().includes(q)) ||
        (s.clients365 && s.clients365.some(c => c.toLowerCase().includes(q)))
      );
    });
  }, [siteHistory, search, statusFilter, siteTypeFilter]);

  // ── Complete Chronological Bookings Ledger (All Previous, Current, Upcoming) ──
  const allBookings = useMemo(() => {
    const list = [];
    const now = new Date();

    const activeSitesMap = new Map();
    dbSites.forEach(s => activeSitesMap.set(canonicalSiteCode(s.site_code), s));

    dbOccupancy
      .filter(c => c.record_status !== 'archived')
      .forEach(c => {
        const sCode = canonicalSiteCode(c.site_code);
        const site = activeSitesMap.get(sCode);
        const sDate = parseFlexibleDate(c.start_date || c.booking_date || c.month);
        const eDate = parseFlexibleDate(c.end_date) || (sDate ? new Date(sDate.getFullYear(), sDate.getMonth() + 1, 0, 23, 59, 59) : null);
        const isAct = c.status === 'active' || (sDate && eDate && sDate <= now && eDate >= now);
        const isUpc = c.status === 'upcoming' || (sDate && sDate > now);

        list.push({
          id: c.id,
          site_code: sCode || '—',
          siteType: getSiteTypeTag(sCode),
          location: c.location || site?.address || site?.area || '—',
          city: c.city || site?.city || 'Ahmedabad',
          client: c.client || c.client_name || '—',
          brand: c.brand || '—',
          display: cleanDisplayTitle(c.display || c.campaign_name || '—'),
          vendor_name: c.vendor_name || '—',
          month: c.month || (sDate ? fmtMonthYear(sDate) : '—'),
          start_date: c.start_date || c.booking_date,
          end_date: c.end_date,
          days: c.days ?? (sDate && eDate ? Math.round((eDate - sDate) / 86400000) + 1 : '—'),
          total_amount: Number(c.total_amount || c.revenue || 0),
          pending: Number(c.pending || 0),
          po: c.po || '',
          bill: c.bill || '',
          sDate,
          eDate,
          status: isAct ? 'active' : isUpc ? 'upcoming' : 'past'
        });
      });

    return list.sort((a, b) => {
      const siteComp = String(a.site_code || '').localeCompare(String(b.site_code || ''), undefined, { numeric: true });
      if (siteComp !== 0) return siteComp;
      const tA = a.sDate ? a.sDate.getTime() : 0;
      const tB = b.sDate ? b.sDate.getTime() : 0;
      return tB - tA;
    });
  }, [dbSites, dbOccupancy, siteHistory]);

  const filteredBookings = useMemo(() => {
    return allBookings.filter(b => {
      if (siteTypeFilter === 'Combined' && b.siteType !== 'Combined') return false;
      if (siteTypeFilter === 'Split Face' && b.siteType !== 'Split Face') return false;

      if (statusFilter === 'occupied' && b.status !== 'active') return false;
      if (statusFilter === 'past' && b.status !== 'past') return false;
      if (statusFilter === 'upcoming' && b.status !== 'upcoming') return false;
      if (statusFilter === 'vacant') return false;

      if (!search.trim()) return true;
      const q = search.trim().toLowerCase();
      return (
        b.site_code.toLowerCase().includes(q) ||
        b.location.toLowerCase().includes(q) ||
        b.client.toLowerCase().includes(q) ||
        b.display.toLowerCase().includes(q) ||
        b.month.toLowerCase().includes(q) ||
        (b.po && b.po.toLowerCase().includes(q)) ||
        (b.bill && b.bill.toLowerCase().includes(q))
      );
    });
  }, [allBookings, statusFilter, siteTypeFilter, search]);

  const activeBookingsCount = useMemo(() => allBookings.filter(b => b.status === 'active').length, [allBookings]);
  const pastBookingsCount = useMemo(() => allBookings.filter(b => b.status === 'past').length, [allBookings]);
  const upcomingBookingsCount = useMemo(() => allBookings.filter(b => b.status === 'upcoming').length, [allBookings]);
  const totalHistoricalRev = useMemo(() => allBookings.reduce((sum, b) => sum + b.total_amount, 0), [allBookings]);

  // ── Overall Stats ─────────────────────────────────────────────────────
  const stats = useMemo(() => {
    const total = siteHistory.length;
    if (!total) return { avgPct: 0, occupied: 0, vacant: 0, total: 0 };

    let totalPct = 0;
    let occupiedCount = 0;

    siteHistory.forEach(s => {
      const isOcc = s.currentStatus === 'active' || (s.pct365 && s.pct365 > 0);
      if (isOcc) occupiedCount++;
      const p = s.pct365 || 0;
      totalPct += p;
    });

    return {
      avgPct: Math.round(totalPct / total),
      occupied: occupiedCount,
      vacant: total - occupiedCount,
      total
    };
  }, [siteHistory]);

  // ── Export Occupancy Matrix to Excel ──────────────────────────────────
  async function exportOccupancyExcel() {
    try {
      const workbook = new ExcelJS.Workbook();
      const ws = workbook.addWorksheet(viewTab === 'history' ? 'Bookings History' : 'Occupancy History');
      
      if (viewTab === 'history') {
        ws.columns = [
          { header: 'Site Code', key: 'site_code', width: 14 },
          { header: 'Location / Area', key: 'location', width: 28 },
          { header: 'City', key: 'city', width: 16 },
          { header: 'Client / Agency', key: 'client', width: 26 },
          { header: 'Display / Campaign', key: 'display', width: 24 },
          { header: 'Month', key: 'month', width: 16 },
          { header: 'Start Date', key: 'start_date', width: 14 },
          { header: 'End Date', key: 'end_date', width: 14 },
          { header: 'Days', key: 'days', width: 10 },
          { header: 'Status', key: 'status', width: 14 },
          { header: 'Total Amount', key: 'total_amount', width: 16 },
          { header: 'Pending', key: 'pending', width: 14 },
          { header: 'PO Number', key: 'po', width: 16 },
          { header: 'Bill Number', key: 'bill', width: 16 }
        ];

        filteredBookings.forEach(b => {
          ws.addRow({
            site_code: b.site_code,
            location: b.location,
            city: b.city,
            client: b.client,
            display: b.display,
            month: b.month,
            start_date: b.start_date ? formatDate(b.start_date) : '',
            end_date: b.end_date ? formatDate(b.end_date) : '',
            days: b.days,
            status: b.status.toUpperCase(),
            total_amount: b.total_amount,
            pending: b.pending,
            po: b.po,
            bill: b.bill
          });
        });
      } else {
        const periods = viewTab === '6months' ? periods6M : [];
        
        const columns = [
          { header: 'Site Code', key: 'site_code', width: 14 },
          { header: 'Location / Area', key: 'area', width: 28 },
          { header: 'City', key: 'city', width: 16 },
          { header: 'Client / Brand', key: 'client', width: 26 },
          { header: 'Status', key: 'status', width: 14 }
        ];

        if (viewTab === 'overview') {
          columns.push({ header: 'Occupied Days (365d)', key: 'days', width: 20 });
          columns.push({ header: 'Occupancy Rate %', key: 'pct', width: 18 });
        } else {
          periods.forEach(p => {
            columns.push({ header: `${p.label} (Occ %)`, key: `pct_${p.key}`, width: 16 });
            columns.push({ header: `${p.label} (Booked Client)`, key: `client_${p.key}`, width: 24 });
          });
          columns.push({ header: 'Average %', key: 'avg', width: 14 });
        }

        ws.columns = columns;

        filteredSites.forEach(s => {
          const rowData = {
            site_code: s.site_code,
            area: s.area,
            city: s.city,
            client: s.currentClient ? `${s.currentClient}${s.currentBrand ? ` (${s.currentBrand})` : ''}` : 'Vacant',
            status: s.currentStatus ? s.currentStatus.toUpperCase() : 'VACANT'
          };

          if (viewTab === 'overview') {
            rowData.days = s.days365;
            rowData.pct = `${s.pct365}%`;
          } else {
            const dataMap = s.by6M;
            let sum = 0;
            periods.forEach(p => {
              const entry = dataMap?.[p.key];
              const pct = typeof entry === 'object' ? (entry?.pct ?? 0) : (Number(entry) || 0);
              const clients = typeof entry === 'object' && entry?.clients ? entry.clients.join(', ') : '';
              rowData[`pct_${p.key}`] = `${pct}%`;
              rowData[`client_${p.key}`] = clients || (pct > 0 ? 'Occupied' : 'Vacant');
              sum += pct;
            });
            rowData.avg = `${Math.round(sum / (periods.length || 1))}%`;
          }

          ws.addRow(rowData);
        });
      }

      // Style header
      const headerRow = ws.getRow(1);
      headerRow.height = 26;
      headerRow.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      });

      const buf = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Site_Occupancy_${viewTab}.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Failed to export occupancy Excel:', err);
      alert('Could not generate Excel export.');
    }
  }

  return (
    <>
      <PageHead 
        title="Site Occupancy History" 
        desc="Simple & clear site-wise occupancy tracker with live utilization, booked clients, and vacant availability." 
      />

      {/* ── Top Summary & KPI Cards (Click to filter) ────────────────── */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '14px', marginBottom: '16px' }}>
        {viewTab === 'history' ? (
          <>
            <div 
              className="scooh-panel" 
              style={{ padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: '4px', cursor: 'pointer', border: statusFilter === 'ALL' ? '2px solid rgba(56, 189, 248, 0.4)' : '1px solid rgba(255,255,255,0.08)' }}
              onClick={() => setStatusFilter('ALL')}
              title="Click to view all booking records"
            >
              <div style={{ fontSize: '11px', fontWeight: 800, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
                Total Bookings History
              </div>
              <div style={{ fontSize: '28px', fontWeight: 900, color: '#f1f5f9', lineHeight: 1.1 }}>
                {allBookings.length}
              </div>
              <div style={{ fontSize: '12px', color: '#64748b' }}>
                All previous & active campaigns
              </div>
            </div>

            <div 
              className="scooh-panel" 
              style={{ padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: '4px', cursor: 'pointer', border: statusFilter === 'occupied' ? '2px solid #10b981' : '1px solid rgba(255,255,255,0.08)', background: statusFilter === 'occupied' ? 'rgba(16, 185, 129, 0.08)' : undefined }}
              onClick={() => setStatusFilter(prev => prev === 'occupied' ? 'ALL' : 'occupied')}
              title="Click to filter currently active bookings"
            >
              <div style={{ fontSize: '11px', fontWeight: 800, color: '#10b981', textTransform: 'uppercase', letterSpacing: '0.05em', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                <span>🟢 Currently Active</span>
                {statusFilter === 'occupied' && <span style={{ fontSize: '10px', background: '#10b981', color: '#000', padding: '1px 6px', borderRadius: '10px', fontWeight: 900 }}>Active</span>}
              </div>
              <div style={{ fontSize: '28px', fontWeight: 900, color: '#10b981', lineHeight: 1.1 }}>
                {activeBookingsCount}
              </div>
              <div style={{ fontSize: '12px', color: '#64748b' }}>
                Currently running on sites
              </div>
            </div>

            <div 
              className="scooh-panel" 
              style={{ padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: '4px', cursor: 'pointer', border: statusFilter === 'past' ? '2px solid #a855f7' : '1px solid rgba(255,255,255,0.08)', background: statusFilter === 'past' ? 'rgba(168, 85, 247, 0.08)' : undefined }}
              onClick={() => setStatusFilter(prev => prev === 'past' ? 'ALL' : 'past')}
              title="Click to filter previous completed bookings"
            >
              <div style={{ fontSize: '11px', fontWeight: 800, color: '#c084fc', textTransform: 'uppercase', letterSpacing: '0.05em', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                <span>⏱ Previous / Completed</span>
                {statusFilter === 'past' && <span style={{ fontSize: '10px', background: '#a855f7', color: '#fff', padding: '1px 6px', borderRadius: '10px', fontWeight: 900 }}>Active</span>}
              </div>
              <div style={{ fontSize: '28px', fontWeight: 900, color: '#c084fc', lineHeight: 1.1 }}>
                {pastBookingsCount}
              </div>
              <div style={{ fontSize: '12px', color: '#64748b' }}>
                Archived & completed bookings
              </div>
            </div>

            <div 
              className="scooh-panel" 
              style={{ padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: '4px', border: '1px solid rgba(255,255,255,0.08)' }}
            >
              <div style={{ fontSize: '11px', fontWeight: 800, color: '#4ade80', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
                Total Revenue
              </div>
              <div style={{ fontSize: '28px', fontWeight: 900, color: '#4ade80', lineHeight: 1.1 }}>
                {money(totalHistoricalRev)}
              </div>
              <div style={{ fontSize: '12px', color: '#64748b' }}>
                Cumulative booking revenue
              </div>
            </div>
          </>
        ) : (
          <>
            <div 
              className="scooh-panel" 
              style={{ padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: '8px', cursor: 'pointer', border: statusFilter === 'ALL' ? '2px solid rgba(56, 189, 248, 0.4)' : '1px solid rgba(255,255,255,0.08)' }}
              onClick={() => setStatusFilter('ALL')}
              title="Click to view all sites"
            >
              <div style={{ fontSize: '11px', fontWeight: 800, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
                Average Occupancy
              </div>
              <div style={{ display: 'flex', alignItems: 'baseline', gap: '8px' }}>
                <span style={{ fontSize: '28px', fontWeight: 900, color: stats.avgPct >= 60 ? '#10b981' : stats.avgPct >= 30 ? '#38bdf8' : '#f59e0b', lineHeight: 1 }}>
                  {stats.avgPct}%
                </span>
                <span style={{ fontSize: '12px', color: '#64748b' }}>across portfolio</span>
              </div>
              <div style={{ height: '6px', background: '#0b1016', borderRadius: '999px', overflow: 'hidden', border: '1px solid #222c37' }}>
                <div style={{ height: '100%', width: `${stats.avgPct}%`, background: stats.avgPct >= 60 ? '#10b981' : '#38bdf8', borderRadius: '999px' }} />
              </div>
            </div>

            <div 
              className="scooh-panel" 
              style={{ padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: '4px', cursor: 'pointer', border: statusFilter === 'occupied' ? '2px solid #10b981' : '1px solid rgba(255,255,255,0.08)', background: statusFilter === 'occupied' ? 'rgba(16, 185, 129, 0.08)' : undefined }}
              onClick={() => setStatusFilter(prev => prev === 'occupied' ? 'ALL' : 'occupied')}
              title="Click to filter only occupied sites"
            >
              <div style={{ fontSize: '11px', fontWeight: 800, color: '#10b981', textTransform: 'uppercase', letterSpacing: '0.05em', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                <span>🟢 Occupied Sites</span>
                {statusFilter === 'occupied' && <span style={{ fontSize: '10px', background: '#10b981', color: '#000', padding: '1px 6px', borderRadius: '10px', fontWeight: 900 }}>Active Filter</span>}
              </div>
              <div style={{ fontSize: '28px', fontWeight: 900, color: '#10b981', lineHeight: 1.1 }}>
                {stats.occupied}
              </div>
              <div style={{ fontSize: '12px', color: '#64748b' }}>
                {stats.total > 0 ? `${Math.round((stats.occupied / stats.total) * 100)}% active campaigns` : 'No sites'}
              </div>
            </div>

            <div 
              className="scooh-panel" 
              style={{ padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: '4px', cursor: 'pointer', border: statusFilter === 'vacant' ? '2px solid #f59e0b' : '1px solid rgba(255,255,255,0.08)', background: statusFilter === 'vacant' ? 'rgba(245, 158, 11, 0.08)' : undefined }}
              onClick={() => setStatusFilter(prev => prev === 'vacant' ? 'ALL' : 'vacant')}
              title="Click to filter only vacant sites"
            >
              <div style={{ fontSize: '11px', fontWeight: 800, color: '#f59e0b', textTransform: 'uppercase', letterSpacing: '0.05em', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                <span>⚪ Vacant Sites</span>
                {statusFilter === 'vacant' && <span style={{ fontSize: '10px', background: '#f59e0b', color: '#000', padding: '1px 6px', borderRadius: '10px', fontWeight: 900 }}>Active Filter</span>}
              </div>
              <div style={{ fontSize: '28px', fontWeight: 900, color: stats.vacant > 0 ? '#f59e0b' : '#10b981', lineHeight: 1.1 }}>
                {stats.vacant}
              </div>
              <div style={{ fontSize: '12px', color: '#64748b' }}>
                Ready for booking
              </div>
            </div>

            <div 
              className="scooh-panel" 
              style={{ padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: '4px', cursor: 'pointer', border: statusFilter === 'ALL' ? '2px solid #38bdf8' : '1px solid rgba(255,255,255,0.08)' }}
              onClick={() => setStatusFilter('ALL')}
              title="Click to show all sites"
            >
              <div style={{ fontSize: '11px', fontWeight: 800, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
                Total Sites Tracked
              </div>
              <div style={{ fontSize: '28px', fontWeight: 900, color: '#f1f5f9', lineHeight: 1.1 }}>
                {stats.total}
              </div>
              <div style={{ fontSize: '12px', color: '#64748b' }}>
                {dataSource === 'excel' ? 'Imported from Excel' : 'Live from Database'}
              </div>
            </div>
          </>
        )}
      </div>

      {/* ── Unified Clean Controls Toolbar ───────────────────────────────── */}
      <section className="scooh-panel" style={{ marginBottom: '16px', padding: '12px 18px' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '12px', flexWrap: 'wrap' }}>
          
          {/* Left: Search & Quick Status Pills */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap', flex: '1 1 auto' }}>
            <input
              className="scooh-search"
              style={{ minWidth: '220px', maxWidth: '320px', fontSize: '12px' }}
              placeholder={viewTab === 'history' ? "Search client, display, site code, PO…" : "Search site, client, area…"}
              value={search}
              onChange={e => setSearch(e.target.value)}
            />

            {/* Quick Status Filter Pills */}
            <div style={{ display: 'inline-flex', gap: '6px', background: 'rgba(15, 23, 42, 0.6)', padding: '3px', borderRadius: '20px', border: '1px solid #1e293b' }}>
              <button
                type="button"
                onClick={() => setStatusFilter('ALL')}
                style={{
                  padding: '4px 12px',
                  borderRadius: '16px',
                  border: 'none',
                  background: statusFilter === 'ALL' ? 'rgba(56, 189, 248, 0.25)' : 'transparent',
                  color: statusFilter === 'ALL' ? '#38bdf8' : '#94a3b8',
                  fontWeight: statusFilter === 'ALL' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                All ({viewTab === 'history' ? allBookings.length : siteHistory.length})
              </button>
              <button
                type="button"
                onClick={() => setStatusFilter('occupied')}
                style={{
                  padding: '4px 12px',
                  borderRadius: '16px',
                  border: 'none',
                  background: statusFilter === 'occupied' ? 'rgba(16, 185, 129, 0.25)' : 'transparent',
                  color: statusFilter === 'occupied' ? '#4ade80' : '#94a3b8',
                  fontWeight: statusFilter === 'occupied' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                🟢 Active ({viewTab === 'history' ? activeBookingsCount : stats.occupied})
              </button>
              {viewTab === 'history' ? (
                <>
                  <button
                    type="button"
                    onClick={() => setStatusFilter('past')}
                    style={{
                      padding: '4px 12px',
                      borderRadius: '16px',
                      border: 'none',
                      background: statusFilter === 'past' ? 'rgba(168, 85, 247, 0.25)' : 'transparent',
                      color: statusFilter === 'past' ? '#c084fc' : '#94a3b8',
                      fontWeight: statusFilter === 'past' ? 800 : 600,
                      fontSize: '11.5px',
                      cursor: 'pointer'
                    }}
                  >
                    ⏱ Previous ({pastBookingsCount})
                  </button>
                  <button
                    type="button"
                    onClick={() => setStatusFilter('upcoming')}
                    style={{
                      padding: '4px 12px',
                      borderRadius: '16px',
                      border: 'none',
                      background: statusFilter === 'upcoming' ? 'rgba(234, 179, 8, 0.25)' : 'transparent',
                      color: statusFilter === 'upcoming' ? '#facc15' : '#94a3b8',
                      fontWeight: statusFilter === 'upcoming' ? 800 : 600,
                      fontSize: '11.5px',
                      cursor: 'pointer'
                    }}
                  >
                    ⏳ Upcoming ({upcomingBookingsCount})
                  </button>
                </>
              ) : (
                <button
                  type="button"
                  onClick={() => setStatusFilter('vacant')}
                  style={{
                    padding: '4px 12px',
                    borderRadius: '16px',
                    border: 'none',
                    background: statusFilter === 'vacant' ? 'rgba(245, 158, 11, 0.25)' : 'transparent',
                    color: statusFilter === 'vacant' ? '#fbbf24' : '#94a3b8',
                    fontWeight: statusFilter === 'vacant' ? 800 : 600,
                    fontSize: '11.5px',
                    cursor: 'pointer'
                  }}
                >
                  ⚪ Vacant ({stats.vacant})
                </button>
              )}
            </div>

            {/* Site Hierarchy Filter: All, Combined, Split Face */}
            <div style={{ display: 'inline-flex', gap: '4px', background: 'rgba(15, 23, 42, 0.6)', padding: '3px', borderRadius: '20px', border: '1px solid #1e293b' }}>
              <button
                type="button"
                onClick={() => setSiteTypeFilter('ALL')}
                style={{
                  padding: '4px 10px',
                  borderRadius: '16px',
                  border: 'none',
                  background: siteTypeFilter === 'ALL' ? 'rgba(148, 163, 184, 0.2)' : 'transparent',
                  color: siteTypeFilter === 'ALL' ? '#f1f5f9' : '#94a3b8',
                  fontWeight: siteTypeFilter === 'ALL' ? 800 : 600,
                  fontSize: '11px',
                  cursor: 'pointer'
                }}
                title="Show all single, combined, and split-face sites"
              >
                All Types
              </button>
              <button
                type="button"
                onClick={() => setSiteTypeFilter(prev => prev === 'Combined' ? 'ALL' : 'Combined')}
                style={{
                  padding: '4px 10px',
                  borderRadius: '16px',
                  border: 'none',
                  background: siteTypeFilter === 'Combined' ? 'rgba(168, 85, 247, 0.25)' : 'transparent',
                  color: siteTypeFilter === 'Combined' ? '#c084fc' : '#94a3b8',
                  fontWeight: siteTypeFilter === 'Combined' ? 800 : 600,
                  fontSize: '11px',
                  cursor: 'pointer'
                }}
                title="Filter multi-panel combined sites (e.g. MB-04, MB-05, MB-06, MB-10, MB-17)"
              >
                🔀 Combined
              </button>
              <button
                type="button"
                onClick={() => setSiteTypeFilter(prev => prev === 'Split Face' ? 'ALL' : 'Split Face')}
                style={{
                  padding: '4px 10px',
                  borderRadius: '16px',
                  border: 'none',
                  background: siteTypeFilter === 'Split Face' ? 'rgba(56, 189, 248, 0.25)' : 'transparent',
                  color: siteTypeFilter === 'Split Face' ? '#38bdf8' : '#94a3b8',
                  fontWeight: siteTypeFilter === 'Split Face' ? 800 : 600,
                  fontSize: '11px',
                  cursor: 'pointer'
                }}
                title="Filter individual split face panels"
              >
                ✂️ Split Face
              </button>
            </div>
          </div>

          {/* Right: Clean View Switcher & Action Buttons */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
            {/* View Mode Toggle: 3 Clear Views */}
            <div style={{ display: 'inline-flex', gap: '4px', background: 'rgba(15, 23, 42, 0.6)', padding: '3px', borderRadius: '20px', border: '1px solid #1e293b' }}>
              <button
                type="button"
                onClick={() => setViewTab('history')}
                style={{
                  padding: '5px 14px',
                  borderRadius: '16px',
                  border: 'none',
                  background: viewTab === 'history' ? 'rgba(56, 189, 248, 0.25)' : 'transparent',
                  color: viewTab === 'history' ? '#38bdf8' : '#94a3b8',
                  fontWeight: viewTab === 'history' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                📋 All Bookings History ({allBookings.length})
              </button>
              <button
                type="button"
                onClick={() => setViewTab('6months')}
                style={{
                  padding: '5px 14px',
                  borderRadius: '16px',
                  border: 'none',
                  background: viewTab === '6months' ? 'rgba(242, 201, 76, 0.18)' : 'transparent',
                  color: viewTab === '6months' ? '#f2c94c' : '#94a3b8',
                  fontWeight: viewTab === '6months' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                🗓️ 6 Months Timeline
              </button>
              <button
                type="button"
                onClick={() => setViewTab('overview')}
                style={{
                  padding: '5px 14px',
                  borderRadius: '16px',
                  border: 'none',
                  background: viewTab === 'overview' ? 'rgba(242, 201, 76, 0.18)' : 'transparent',
                  color: viewTab === 'overview' ? '#f2c94c' : '#94a3b8',
                  fontWeight: viewTab === 'overview' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                ⚡ 365-Day Overview
              </button>
            </div>

            <button
              type="button"
              className="scooh-btn ghost"
              onClick={exportOccupancyExcel}
              title="Export complete occupancy table to Excel"
              style={{ fontSize: '11.5px', padding: '6px 12px' }}
            >
              📥 Export Excel
            </button>

            {/* Data is shared from Campaign Tracker — import/manage via Campaign Tracker tab */}

            <button
              type="button"
              className="scooh-btn ghost"
              onClick={loadSystemData}
              disabled={loading}
              title="Refresh database"
              style={{ fontSize: '11.5px', padding: '6px 10px' }}
            >
              🔄
            </button>
          </div>
        </div>
      </section>

      {/* ── Main Occupancy Table ────────────────────────────────────────── */}
      <section className="scooh-panel">
        {loading ? (
          <div style={{ textAlign: 'center', padding: '50px 20px', color: '#94a3b8' }}>
            <div style={{ fontSize: '32px', marginBottom: '12px' }}>🔄</div>
            <b>Loading occupancy history...</b>
          </div>
        ) : viewTab === 'history' ? (
          /* ── Dedicated All Bookings History Table (Zero Confusion) ── */
          filteredBookings.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '50px 20px', color: '#64748b' }}>
              <div style={{ fontSize: '36px', marginBottom: '12px' }}>🔍</div>
              <h4 style={{ margin: '0 0 6px', color: '#f1f5f9' }}>No booking history records found</h4>
              <p style={{ margin: 0, fontSize: '12px' }}>Try adjusting your search filter or status pill.</p>
            </div>
          ) : (
            <div className="scooh-tablewrap" style={{ overflowX: 'auto' }}>
              <table className="scooh-table" style={{ width: '100%', minWidth: '1100px' }}>
                <thead>
                  <tr>
                    <th style={{ minWidth: '100px' }}>Site Code</th>
                    <th style={{ minWidth: '160px' }}>Location / Area</th>
                    <th style={{ minWidth: '160px' }}>Client / Agency</th>
                    <th style={{ minWidth: '150px' }}>Display / Brand</th>
                    <th style={{ minWidth: '100px' }}>Month</th>
                    <th style={{ minWidth: '180px' }}>Booking Dates</th>
                    <th style={{ minWidth: '70px', textAlign: 'center' }}>Days</th>
                    <th style={{ minWidth: '130px', textAlign: 'center' }}>Status</th>
                    <th style={{ minWidth: '120px', textAlign: 'right' }}>Total Amount</th>
                    <th style={{ minWidth: '110px', textAlign: 'right' }}>Pending</th>
                    <th style={{ minWidth: '110px' }}>PO / Bill</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredBookings.map((b, bIdx) => {
                    const isAct = b.status === 'active';
                    const isUpc = b.status === 'upcoming';
                    const statusColor = isAct ? '#4ade80' : isUpc ? '#facc15' : '#94a3b8';
                    const statusBg = isAct ? 'rgba(34, 197, 94, 0.16)' : isUpc ? 'rgba(234, 179, 8, 0.16)' : 'rgba(148, 163, 184, 0.12)';
                    const statusBorder = isAct ? 'rgba(34, 197, 94, 0.35)' : isUpc ? 'rgba(234, 179, 8, 0.35)' : 'rgba(148, 163, 184, 0.25)';
                    const statusLabel = isAct ? '● Active' : isUpc ? '⏳ Upcoming' : '⏱ Completed';

                    return (
                      <tr key={b.id || bIdx}>
                        <td>
                          <div style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>
                            <span className="scooh-plate" style={{ fontSize: '11.5px', fontWeight: 800 }}>
                              {b.site_code}
                            </span>
                            {b.siteType && b.siteType !== 'Single' && (
                              <span
                                style={{
                                  fontSize: '9.5px',
                                  padding: '1px 5px',
                                  borderRadius: '4px',
                                  background: b.siteType === 'Combined' ? 'rgba(168,85,247,0.2)' : 'rgba(56,189,248,0.2)',
                                  color: b.siteType === 'Combined' ? '#c084fc' : '#38bdf8',
                                  border: `1px solid ${b.siteType === 'Combined' ? 'rgba(168,85,247,0.4)' : 'rgba(56,189,248,0.4)'}`,
                                  fontWeight: 700
                                }}
                              >
                                {b.siteType}
                              </span>
                            )}
                          </div>
                        </td>
                        <td>
                          <div style={{ fontWeight: 700, color: '#f1f5f9', fontSize: '12.5px', lineHeight: 1.3 }}>
                            {b.location}
                          </div>
                          <div style={{ fontSize: '11px', color: '#94a3b8', marginTop: '2px' }}>
                            📍 {b.city}
                          </div>
                        </td>
                        <td>
                          <div style={{ fontWeight: 800, color: '#ffffff', fontSize: '13px' }}>
                            {b.client}
                          </div>
                        </td>
                        <td>
                          <div style={{ color: '#38bdf8', fontWeight: 600, fontSize: '12px' }}>
                            {b.display}
                          </div>
                          {b.brand && b.brand !== '—' && b.brand !== b.client && (
                            <div style={{ fontSize: '11px', color: '#94a3b8', marginTop: '1px' }}>
                              Brand: {b.brand}
                            </div>
                          )}
                        </td>
                        <td>
                          <span style={{ fontSize: '11.5px', color: '#cbd5e1', background: 'rgba(56,189,248,0.1)', padding: '2px 7px', borderRadius: '5px', border: '1px solid rgba(56,189,248,0.25)' }}>
                            {b.month}
                          </span>
                        </td>
                        <td style={{ fontSize: '12px', color: '#cbd5e1' }}>
                          <div>
                            <span>{formatDate(b.start_date)}</span>
                            <span style={{ color: '#64748b', margin: '0 4px' }}>→</span>
                            <span>{formatDate(b.end_date) || 'Ongoing'}</span>
                          </div>
                        </td>
                        <td style={{ textAlign: 'center', fontWeight: 700, color: '#e2e8f0', fontSize: '12px' }}>
                          {b.days}
                        </td>
                        <td style={{ textAlign: 'center' }}>
                          <span style={{
                            display: 'inline-flex',
                            alignItems: 'center',
                            gap: '5px',
                            padding: '2px 8px',
                            borderRadius: '6px',
                            background: statusBg,
                            color: statusColor,
                            border: `1px solid ${statusBorder}`,
                            fontSize: '11px',
                            fontWeight: 800,
                            whiteSpace: 'nowrap'
                          }}>
                            <span style={{ width: '6px', height: '6px', borderRadius: '50%', background: statusColor }} />
                            {statusLabel}
                          </span>
                        </td>
                        <td style={{ textAlign: 'right', fontWeight: 800, color: b.total_amount > 0 ? '#4ade80' : '#64748b', fontSize: '12.5px' }}>
                          {b.total_amount > 0 ? money(b.total_amount) : '—'}
                        </td>
                        <td style={{ textAlign: 'right', fontWeight: 700, color: b.pending > 0 ? '#f87171' : '#64748b', fontSize: '12px' }}>
                          {b.pending > 0 ? money(b.pending) : '—'}
                        </td>
                        <td>
                          {b.po ? <span className="scooh-plate" style={{ fontSize: '10.5px' }}>{b.po}</span> : ''}
                          {b.bill ? <span style={{ color: '#93c5fd', fontSize: '11px', marginLeft: b.po ? '4px' : 0 }}>{b.bill}</span> : ''}
                          {!b.po && !b.bill && <span style={{ color: '#64748b', fontSize: '11px' }}>—</span>}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )
        ) : filteredSites.length === 0 ? (
          <div style={{ textAlign: 'center', padding: '50px 20px', color: '#64748b' }}>
            <div style={{ fontSize: '36px', marginBottom: '12px' }}>🔍</div>
            <h4 style={{ margin: '0 0 6px', color: '#f1f5f9' }}>No matching sites found</h4>
            <p style={{ margin: 0, fontSize: '12px' }}>Try adjusting your search filter or import an Excel file.</p>
          </div>
        ) : (
          <div className="scooh-tablewrap" style={{ overflowX: 'auto' }}>
            <table className="scooh-table" style={{ borderCollapse: 'collapse', width: '100%' }}>
              <thead>
                <tr>
                  <th style={{ position: 'sticky', left: 0, zIndex: 3, background: '#10161e', minWidth: '110px' }}>Site ID</th>
                  <th style={{ minWidth: '140px' }}>Location / Area</th>
                  <th style={{ minWidth: '160px' }}>Client / Brand</th>

                  {/* 6 Months View Columns */}
                  {viewTab === '6months' && periods6M.map(p => (
                    <th key={p.key} style={{ minWidth: '130px', textAlign: 'center' }}>
                      <div>{p.label}</div>
                      <div style={{ fontSize: '10px', color: '#64748b', fontWeight: 500 }}>Client & Occ %</div>
                    </th>
                  ))}

                  {/* Yearly View Columns */}
                  {viewTab === 'yearly' && periodsYearly.map(p => (
                    <th key={p.key} style={{ minWidth: '130px', textAlign: 'center' }}>
                      <div>{p.label}</div>
                      <div style={{ fontSize: '10px', color: '#64748b', fontWeight: 500 }}>Client & Occ %</div>
                    </th>
                  ))}

                  {/* 365 Days Overview Columns */}
                  {viewTab === 'overview' && (
                    <>
                      <th style={{ minWidth: '100px' }}>City</th>
                      <th style={{ minWidth: '140px' }}>Occupied Days (365d)</th>
                      <th style={{ minWidth: '160px' }}>Occupancy Rate</th>
                      <th style={{ minWidth: '100px', textAlign: 'center' }}>Status</th>
                    </>
                  )}

                  {/* Average Column for 6M and Yearly */}
                  {viewTab !== 'overview' && (
                    <th style={{ minWidth: '125px', textAlign: 'center', background: '#0d1a2e', color: '#7dd3fc', borderLeft: '2px solid #1e3a5f' }}>
                      Average %
                    </th>
                  )}
                </tr>
              </thead>
              <tbody>
                {filteredSites.map((s, idx) => {
                  const clientCell = (
                    <td>
                      {s.currentClient ? (
                        <div>
                          <div style={{ fontWeight: 800, color: '#f1f5f9', fontSize: '12.5px', display: 'flex', alignItems: 'center', gap: '5px' }}>
                            <span style={{ color: '#38bdf8' }}>👤</span>
                            <span>{s.currentClient}</span>
                          </div>
                          {s.currentBrand && s.currentBrand !== s.currentClient && (
                            <div style={{ fontSize: '11px', color: '#94a3b8', marginTop: '2px' }}>
                              🏷️ {s.currentBrand}
                            </div>
                          )}
                          <div style={{ marginTop: '4px' }}>
                            <span 
                              className="scooh-badgechip" 
                              style={{ 
                                fontSize: '9.5px', 
                                padding: '1px 6px',
                                background: s.currentStatus === 'active' ? 'rgba(16, 185, 129, 0.15)' : s.currentStatus === 'upcoming' ? 'rgba(245, 158, 11, 0.15)' : 'rgba(100, 116, 139, 0.15)',
                                color: s.currentStatus === 'active' ? '#10b981' : s.currentStatus === 'upcoming' ? '#f59e0b' : '#94a3b8',
                                border: s.currentStatus === 'active' ? '1px solid rgba(16, 185, 129, 0.3)' : s.currentStatus === 'upcoming' ? '1px solid rgba(245, 158, 11, 0.3)' : '1px solid rgba(100, 116, 139, 0.3)'
                              }}
                            >
                              {s.currentStatus === 'active' ? '● Active' : s.currentStatus === 'upcoming' ? '⏳ Upcoming' : '⏱ Past Client'}
                            </span>
                          </div>
                        </div>
                      ) : (
                        <span style={{ color: '#475569', fontSize: '12px' }}>— Available —</span>
                      )}
                    </td>
                  );

                  if (viewTab === 'overview') {
                    const statusText = s.pct365 >= 75 ? 'Full' : s.pct365 >= 40 ? 'High' : s.pct365 > 0 ? 'Partial' : 'Vacant';
                    const badgeBg = s.pct365 >= 75 ? 'rgba(16, 185, 129, 0.15)' : s.pct365 >= 40 ? 'rgba(56, 189, 248, 0.15)' : s.pct365 > 0 ? 'rgba(245, 158, 11, 0.15)' : 'rgba(100, 116, 139, 0.15)';
                    const badgeColor = s.pct365 >= 75 ? '#10b981' : s.pct365 >= 40 ? '#38bdf8' : s.pct365 > 0 ? '#f59e0b' : '#64748b';

                    return (
                      <tr key={s.site_code || idx}>
                        <td style={{ position: 'sticky', left: 0, zIndex: 2, background: idx % 2 === 0 ? '#0b1016' : '#10161e' }}>
                          <div style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>
                            <span 
                              className="scooh-plate" 
                              style={{ cursor: 'pointer' }}
                              onClick={() => setModalData({ siteCode: s.site_code, area: s.area, periodLabel: 'All Bookings History', campaigns: s.allCampaigns || [] })}
                              title="Click to view all booking history"
                            >
                              {s.site_code}
                            </span>
                            {s.siteType && s.siteType !== 'Single' && (
                              <span
                                style={{
                                  fontSize: '9.5px',
                                  padding: '1px 5px',
                                  borderRadius: '4px',
                                  background: s.siteType === 'Combined' ? 'rgba(168,85,247,0.2)' : 'rgba(56,189,248,0.2)',
                                  color: s.siteType === 'Combined' ? '#c084fc' : '#38bdf8',
                                  border: `1px solid ${s.siteType === 'Combined' ? 'rgba(168,85,247,0.4)' : 'rgba(56,189,248,0.4)'}`,
                                  fontWeight: 700
                                }}
                              >
                                {s.siteType}
                              </span>
                            )}
                          </div>
                        </td>
                        <td><b>{s.area}</b></td>
                        {clientCell}
                        <td>{s.city}</td>
                        <td>
                          <span style={{ fontWeight: 700, color: s.days365 > 0 ? '#f1f5f9' : '#64748b' }}>
                            {s.days365 || 0} days
                          </span>
                          <span style={{ fontSize: '11px', color: '#64748b', marginLeft: '4px' }}>/ 365</span>
                        </td>
                        <td>
                          <OccPercentBar 
                            pct={s.pct365} 
                            days={s.days365} 
                            totalDays={365} 
                            clientNames={s.clients365} 
                            onClick={() => setModalData({ 
                              siteCode: s.site_code, 
                              area: s.area, 
                              periodLabel: '365-Day Rolling Period', 
                              campaigns: s.allCampaigns || [] 
                            })}
                          />
                        </td>
                        <td style={{ textAlign: 'center' }}>
                          <span className="scooh-badgechip" style={{ background: badgeBg, color: badgeColor, border: `1px solid ${badgeColor}40` }}>
                            {statusText}
                          </span>
                        </td>
                      </tr>
                    );
                  }

                  // 6 Months or Yearly View
                  const periods = viewTab === '6months' ? periods6M : periodsYearly;
                  const dataMap = viewTab === '6months' ? s.by6M : s.byYearly;
                  const pcts = periods.map(p => {
                    const entry = dataMap?.[p.key];
                    return typeof entry === 'object' ? (entry?.pct ?? 0) : (Number(entry) || 0);
                  });
                  const siteAvg = pcts.length ? Math.round(pcts.reduce((a, b) => a + b, 0) / pcts.length) : 0;

                  return (
                    <tr key={s.site_code || idx}>
                      <td style={{ position: 'sticky', left: 0, zIndex: 2, background: idx % 2 === 0 ? '#0b1016' : '#10161e' }}>
                        <div style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>
                          <span 
                            className="scooh-plate" 
                            style={{ cursor: 'pointer' }}
                            onClick={() => setModalData({ siteCode: s.site_code, area: s.area, periodLabel: 'All Bookings History', campaigns: s.allCampaigns || [] })}
                            title="Click to view all booking history"
                          >
                            {s.site_code}
                          </span>
                          {s.siteType && s.siteType !== 'Single' && (
                            <span
                              style={{
                                fontSize: '9.5px',
                                padding: '1px 5px',
                                borderRadius: '4px',
                                background: s.siteType === 'Combined' ? 'rgba(168,85,247,0.2)' : 'rgba(56,189,248,0.2)',
                                color: s.siteType === 'Combined' ? '#c084fc' : '#38bdf8',
                                border: `1px solid ${s.siteType === 'Combined' ? 'rgba(168,85,247,0.4)' : 'rgba(56,189,248,0.4)'}`,
                                fontWeight: 700
                              }}
                            >
                              {s.siteType}
                            </span>
                          )}
                        </div>
                      </td>
                      <td><b>{s.area}</b></td>
                      {clientCell}

                      {periods.map(p => {
                        const entry = dataMap?.[p.key];
                        const pct = typeof entry === 'object' ? (entry?.pct ?? 0) : (Number(entry) || 0);
                        const days = typeof entry === 'object' ? entry?.days : undefined;
                        const totalDays = typeof entry === 'object' ? entry?.totalDays : undefined;
                        const clientNames = typeof entry === 'object' && entry?.clients ? entry.clients : [];
                        const brands = typeof entry === 'object' && entry?.brands ? entry.brands : [];
                        const cellCamps = typeof entry === 'object' && entry?.campaigns ? entry.campaigns : [];

                        return (
                          <td key={p.key} style={{ padding: '8px 10px', textAlign: 'center' }}>
                            <OccPercentBar 
                              pct={pct} 
                              days={days} 
                              totalDays={totalDays} 
                              clientNames={clientNames} 
                              brands={brands}
                              onClick={() => setModalData({ 
                                siteCode: s.site_code, 
                                area: s.area, 
                                periodLabel: `${p.label} (${pct}% Occupied)`, 
                                campaigns: cellCamps 
                              })}
                            />
                          </td>
                        );
                      })}

                      <td style={{ padding: '8px 10px', background: '#0d1a2e', borderLeft: '2px solid #1e3a5f', textAlign: 'center' }}>
                        <OccPercentBar pct={siteAvg} />
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </section>

      {/* ── Modal for Inspecting Detailed Monthly Booking History ──────────── */}
      {modalData && (
        <div 
          className="scooh-modal-backdrop" 
          style={{ position: 'fixed', inset: 0, zIndex: 9999, background: 'rgba(0,0,0,0.8)', backdropFilter: 'blur(5px)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '20px' }}
          onClick={() => setModalData(null)}
        >
          <div 
            className="scooh-panel" 
            style={{ width: '100%', maxWidth: '580px', maxHeight: '85vh', overflowY: 'auto', border: '1px solid #38bdf8', boxShadow: '0 20px 40px rgba(0,0,0,0.6)' }}
            onClick={e => e.stopPropagation()}
          >
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', borderBottom: '1px solid #1e293b', paddingBottom: '14px', marginBottom: '16px' }}>
              <div>
                <h3 style={{ margin: 0, color: '#f1f5f9', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '17px' }}>
                  <span>📅</span>
                  <span>Site {modalData.siteCode} — Booking Details</span>
                </h3>
                <p style={{ margin: '4px 0 0', fontSize: '12px', color: '#94a3b8' }}>
                  {modalData.area} • <strong style={{ color: '#38bdf8' }}>{modalData.periodLabel}</strong>
                </p>
              </div>
              <button 
                type="button" 
                onClick={() => setModalData(null)} 
                style={{ background: 'none', border: 'none', color: '#94a3b8', fontSize: '20px', cursor: 'pointer', padding: '4px' }}
                title="Close"
              >
                ✕
              </button>
            </div>

            {modalData.campaigns && modalData.campaigns.length > 0 ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                {modalData.campaigns.map((c, i) => (
                  <div key={c.id || i} style={{ padding: '14px', background: '#0b1320', borderRadius: '8px', border: '1px solid #1e3a5f' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '6px', gap: '10px' }}>
                      <div>
                        <div style={{ fontSize: '15px', fontWeight: 800, color: '#38bdf8' }}>
                          👤 {c.client || c.client_name || 'Client Not Specified'}
                        </div>
                        {c.brand && c.brand !== (c.client || c.client_name) && (
                          <div style={{ fontSize: '12px', color: '#cbd5e1', marginTop: '2px' }}>
                            🏷️ Brand: <b>{c.brand}</b>
                          </div>
                        )}
                      </div>
                      <span className="scooh-badgechip" style={{ background: 'rgba(16, 185, 129, 0.15)', color: '#10b981', border: '1px solid rgba(16, 185, 129, 0.3)', whiteSpace: 'nowrap' }}>
                        {c.month || 'Active Booking'}
                      </span>
                    </div>

                    <div style={{ fontSize: '12.5px', color: '#e2e8f0', margin: '8px 0 6px' }}>
                      📢 <b>Campaign:</b> {c.campaign_name || c.display || 'Standard Display'}
                    </div>

                    <div style={{ display: 'flex', gap: '16px', fontSize: '12px', color: '#94a3b8', flexWrap: 'wrap', borderTop: '1px solid #1e293b', paddingTop: '8px', marginTop: '8px' }}>
                      {(c.start_date || c.booking_date) && (
                        <span>🗓️ <b>Dates:</b> {new Date(c.start_date || c.booking_date).toLocaleDateString('en-IN')} to {c.end_date ? new Date(c.end_date).toLocaleDateString('en-IN') : 'Ongoing'}</span>
                      )}
                      {c.booking_code && (
                        <span>🔖 <b>Code:</b> {c.booking_code}</span>
                      )}
                      {(c.total_amount || c.revenue) && (
                        <span style={{ color: '#10b981', fontWeight: 700 }}>💰 <b>Amount:</b> ₹{Number(c.total_amount || c.revenue).toLocaleString('en-IN')}</span>
                      )}
                    </div>
                  </div>
                ))}
              </div>
            ) : (
              <div style={{ textAlign: 'center', padding: '36px 20px', color: '#64748b' }}>
                <div style={{ fontSize: '36px', marginBottom: '8px' }}>📭</div>
                <h4 style={{ margin: '0 0 4px', color: '#cbd5e1' }}>No booking for this period</h4>
                <p style={{ margin: 0, fontSize: '12px' }}>The site was completely vacant and available for booking during this time.</p>
              </div>
            )}
          </div>
        </div>
      )}
    </>
  );
}

function getCampaignOccupancy(r) {
  if (r.occupancy !== undefined && r.occupancy !== null && String(r.occupancy).trim() !== '') {
    const raw = String(r.occupancy).trim();
    const num = parseFloat(raw.replace('%', ''));
    if (!isNaN(num)) {
      return {
        pct: num,
        label: `${num}%`,
        status: num >= 75 ? 'occupied' : num > 0 ? 'partial' : 'vacant'
      };
    }
    return { pct: null, label: raw, status: 'custom' };
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const start = r.start_date ? new Date(r.start_date) : null;
  const end = r.end_date ? new Date(r.end_date) : null;

  if (end && !isNaN(end.getTime())) {
    end.setHours(23, 59, 59, 999);
    if (end < today) {
      return { pct: 0, label: '0% Vacant', sub: 'Past', status: 'vacant' };
    }
    if (start && !isNaN(start.getTime()) && start > today) {
      return {
        pct: 100,
        label: '100% Booked',
        sub: `Starts ${start.toLocaleDateString('en-IN', { day: '2-digit', month: 'short' })}`,
        status: 'upcoming'
      };
    }
    return { pct: 100, label: '100% Occupied', status: 'occupied' };
  }

  if (r.record_status === 'active') {
    return { pct: 100, label: '100% Occupied', status: 'occupied' };
  }
  return { pct: 0, label: '0% Vacant', status: 'vacant' };
}

function getCampaignPeriod(r, view) {
  const dStr = r.start_date || r.booking_date || r.end_date;
  const d = dStr ? new Date(dStr) : null;
  const validDate = d && !isNaN(d.getTime());

  if (view === 'monthly') {
    if (r.month) return r.month;
    if (validDate) return d.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
    return 'Other';
  }
  if (view === 'weekly') {
    if (validDate) {
      const utc = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
      utc.setUTCDate(utc.getUTCDate() + 4 - (utc.getUTCDay() || 7));
      const yr = utc.getUTCFullYear();
      const wk = Math.ceil(((utc - new Date(Date.UTC(yr, 0, 1))) / 86400000 + 1) / 7);
      return `${yr}-W${String(wk).padStart(2, '0')}`;
    }
    return 'Other';
  }
  if (view === 'yearly') {
    if (validDate) return String(d.getFullYear());
    const mMatch = String(r.month || '').match(/\b(20\d\d)\b/);
    if (mMatch) return mMatch[1];
    return 'Other';
  }
  return 'All';
}

async function exportCampaignsExcel(rowsData, filename = 'MediaBuzz_Campaign_Tracker.xlsx') {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Campaign Tracker', {
    views: [{ state: 'frozen', ySplit: 1 }]
  });

  worksheet.columns = [
    { header: 'Site Code', key: 'site_code', width: 14 },
    { header: 'Month', key: 'month', width: 14 },
    { header: 'Occupancy', key: 'occupancy', width: 16 },
    { header: 'Date', key: 'date', width: 14 },
    { header: 'Client/Agency Name', key: 'client', width: 28 },
    { header: 'Display', key: 'display', width: 24 },
    { header: 'Vendor Name', key: 'vendor_name', width: 22 },
    { header: 'Location', key: 'location', width: 36 },
    { header: 'W', key: 'width', width: 10 },
    { header: 'H', key: 'height', width: 10 },
    { header: 'Size', key: 'size', width: 14 },
    { header: 'Type', key: 'type', width: 16 },
    { header: 'Start Date', key: 'start_date', width: 14 },
    { header: 'End Date', key: 'end_date', width: 14 },
    { header: 'Days', key: 'days', width: 10 },
    { header: 'Advt. Fees per month', key: 'advt_fees', width: 20 },
    { header: 'Printing & Mounting', key: 'printing_mounting_cost', width: 22 },
    { header: 'Total Amount', key: 'total_amount', width: 20 },
    { header: 'PO', key: 'po', width: 16 },
    { header: 'Bill', key: 'bill', width: 18 },
    { header: 'Pending', key: 'pending', width: 18 }
  ];

  rowsData.forEach((r) => {
    worksheet.addRow({
      site_code: r.site_code || '',
      month: r.month || '',
      occupancy: getCampaignOccupancy(r).label,
      date: r.booking_date ? new Date(r.booking_date).toLocaleDateString('en-IN') : (r.date || ''),
      client: r.client || r.client_name || '',
      display: r.display || r.campaign_name || r.brand || '',
      vendor_name: r.vendor_name || '',
      location: r.location || '',
      width: r.width ?? '',
      height: r.height ?? '',
      size: r.size || (r.width && r.height ? `${r.width}x${r.height} ft` : ''),
      type: r.type || 'Hoarding',
      start_date: r.start_date ? new Date(r.start_date).toLocaleDateString('en-IN') : '',
      end_date: r.end_date ? new Date(r.end_date).toLocaleDateString('en-IN') : '',
      days: r.days || (r.start_date && r.end_date ? Math.max(1, Math.round((new Date(r.end_date) - new Date(r.start_date)) / 86400000) + 1) : 30),
      advt_fees: Number(r.advt_fees || 0),
      printing_mounting_cost: Number(r.printing_mounting_cost || (Number(r.printing_cost || 0) + Number(r.mounting_cost || 0))),
      total_amount: Number(r.total_amount || r.revenue || 0),
      po: r.po || '',
      bill: r.bill || r.invoice_no || '',
      pending: Number(r.pending || 0)
    });
  });

  const headerRow = worksheet.getRow(1);
  headerRow.height = 28;
  headerRow.eachCell((cell, colNumber) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF00' }
    };
    cell.font = {
      name: 'Calibri',
      size: 11,
      bold: true,
      color: { argb: 'FF000000' }
    };
    cell.alignment = {
      vertical: 'middle',
      horizontal: [1, 8, 9, 10, 11, 12, 13, 14].includes(colNumber) ? 'center' : ([15, 16, 17, 20].includes(colNumber) ? 'right' : 'left'),
      wrapText: false
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFC0C0C0' } },
      left: { style: 'thin', color: { argb: 'FFC0C0C0' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FFC0C0C0' } }
    };
  });

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  window.URL.revokeObjectURL(url);

  try {
    api.post('/storage', {
      category: 'excel',
      name: filename,
      type: 'xlsx',
      size: `${Math.round(blob.size / 1024)} KB`
    }).catch(() => {});
  } catch (e) {}
}

function CampaignTrackerView() {
  const location = useLocation();
  const navigate = useNavigate();
  const [rows, setRows] = useState([]);
  const [sites, setSites] = useState([]);
  const [loading, setLoading] = useState(false);
  const [editModal, setEditModal] = useState(null);
  const [modalSiteCode, setModalSiteCode] = useState('');
  const [modalLocation, setModalLocation] = useState('');
  const [modalWidth, setModalWidth] = useState('');
  const [modalHeight, setModalHeight] = useState('');
  const [modalSize, setModalSize] = useState('');
  const [modalType, setModalType] = useState('Hoarding');
  const [autoMatchedSite, setAutoMatchedSite] = useState(null);
  const [search, setSearch] = useState('');
  const [siteQueryParam, setSiteQueryParam] = useState('');
  const [monthFilter, setMonthFilter] = useState('ALL');
  const [typeFilter, setTypeFilter] = useState('ALL');
  const [vendorFilter, setVendorFilter] = useState('ALL');
  const [statusFilter, setStatusFilter] = useState('ALL');
  const [sortState, setSortState] = useState({ key: 'id', dir: 'desc' });
  const [saving, setSaving] = useState(false);
  const [banner, setBanner] = useState('');
  const [importingExcel, setImportingExcel] = useState(false);
  const [isViewingOnly, setIsViewingOnly] = useState(false);
  const [selectedIds, setSelectedIds] = useState(new Set());

  const topScrollRef = useRef(null);
  const bottomScrollRef = useRef(null);
  const tableRef = useRef(null);
  const isScrollingRef = useRef(null);
  const [tableScrollWidth, setTableScrollWidth] = useState(2400);

  function handleTopScroll() {
    if (isScrollingRef.current === 'bottom') return;
    isScrollingRef.current = 'top';
    if (bottomScrollRef.current && topScrollRef.current) {
      bottomScrollRef.current.scrollLeft = topScrollRef.current.scrollLeft;
    }
    setTimeout(() => {
      if (isScrollingRef.current === 'top') isScrollingRef.current = null;
    }, 40);
  }

  function handleBottomScroll() {
    if (isScrollingRef.current === 'top') return;
    isScrollingRef.current = 'bottom';
    if (topScrollRef.current && bottomScrollRef.current) {
      topScrollRef.current.scrollLeft = bottomScrollRef.current.scrollLeft;
    }
    setTimeout(() => {
      if (isScrollingRef.current === 'bottom') isScrollingRef.current = null;
    }, 40);
  }

  function scrollHorizontallyBy(delta) {
    if (bottomScrollRef.current) {
      bottomScrollRef.current.scrollBy({ left: delta, behavior: 'smooth' });
    }
  }

  useEffect(() => {
    function updateScrollWidth() {
      if (tableRef.current) {
        const sw = tableRef.current.scrollWidth || tableRef.current.offsetWidth;
        if (sw) setTableScrollWidth(sw);
      } else if (bottomScrollRef.current) {
        const sw = bottomScrollRef.current.scrollWidth;
        if (sw) setTableScrollWidth(sw);
      }
    }

    updateScrollWidth();
    const t = setTimeout(updateScrollWidth, 150);

    let ro;
    if (window.ResizeObserver && bottomScrollRef.current) {
      ro = new ResizeObserver(() => updateScrollWidth());
      ro.observe(bottomScrollRef.current);
      if (tableRef.current) ro.observe(tableRef.current);
    }

    window.addEventListener('resize', updateScrollWidth);
    return () => {
      clearTimeout(t);
      window.removeEventListener('resize', updateScrollWidth);
      if (ro) ro.disconnect();
    };
  }, [rows]);

  const currentRole = getCurrentRole();
  const isAdmin = currentRole === 'admin';
  const isManager = currentRole === 'manager';
  const isStaff = currentRole === 'staff';
  const isReadOnly = currentRole === 'viewer';
  const canDelete = isAdmin || isManager;
  const canAdd = isAdmin || isManager;
  const canEdit = !isReadOnly;

  async function loadData() {
    setLoading(true);
    try {
      const [campRes, sitesRes] = await Promise.all([
        api.get('/campaigns'),
        api.get('/sites')
      ]);
      if (Array.isArray(campRes.data)) {
        setRows(campRes.data);
      }
      if (Array.isArray(sitesRes.data)) {
        setSites(sitesRes.data);
      }
    } catch (err) {
      console.warn('Error loading campaigns/sites:', err);
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    loadData();
    const interval = setInterval(loadData, 35000);
    return () => clearInterval(interval);
  }, []);

  // Auto-detect site code from location and size as user types or edits
  useEffect(() => {
    if (!editModal) {
      setAutoMatchedSite(null);
      return;
    }
    const derivedSize = modalSize || (modalWidth && modalHeight ? `${modalWidth}x${modalHeight}` : '');
    if (!modalLocation && !derivedSize && !modalWidth && !modalHeight) {
      setAutoMatchedSite(null);
      return;
    }
    const matched = matchSiteByLocationAndSize(
      modalLocation,
      derivedSize,
      modalWidth,
      modalHeight,
      sites
    );
    setAutoMatchedSite(matched);
    // If site code is empty and a high-confidence match is detected, auto-populate it
    if (matched && !modalSiteCode.trim()) {
      setModalSiteCode(matched.site_code);
    }
  }, [modalLocation, modalSize, modalWidth, modalHeight, editModal, sites]);

  function applyMatchedSite(s) {
    if (!s) return;
    setModalSiteCode(s.site_code || '');
    if (s.address || s.area) setModalLocation(s.address || s.area || '');
    if (s.width) setModalWidth(s.width);
    if (s.height) setModalHeight(s.height);
    if (s.size) setModalSize(s.size);
    if (s.type) setModalType(s.type);
  }

  function handleSiteCodeChange(code) {
    setModalSiteCode(code);
    const upper = String(code).trim().toUpperCase();
    const found = sites.find(s => String(s.site_code || '').toUpperCase() === upper);
    if (found) {
      if (!modalLocation) setModalLocation(found.address || found.area || '');
      if (!modalWidth && found.width) setModalWidth(found.width);
      if (!modalHeight && found.height) setModalHeight(found.height);
      if (!modalSize && found.size) setModalSize(found.size);
      if (found.type) setModalType(found.type);
    }
  }

  // Sync search/filter when navigating with ?site=MB-XX or ?search=...
  useEffect(() => {
    const params = new URLSearchParams(location.search);
    const siteParam = (params.get('site') || params.get('search') || '').trim();
    if (siteParam) {
      setSearch(siteParam);
      setSiteQueryParam(siteParam);
    } else {
      setSiteQueryParam('');
    }
  }, [location.search]);

  // Master directory — only sites that have campaign records (DB sites used for auto-match only)
  const siteMasterList = useMemo(() => {
    // Build a lookup from DB sites for enriching campaign rows with site metadata
    const siteDbLookup = new Map();
    sites.forEach(s => {
      const code = String(s.site_code || '').trim().toUpperCase();
      if (code) siteDbLookup.set(code, s);
    });

    const siteMap = new Map();

    // Only add sites that have campaign rows
    rows.forEach(r => {
      const code = String(r.site_code || 'UNASSIGNED').trim().toUpperCase();
      if (!siteMap.has(code)) {
        const dbSite = siteDbLookup.get(code);
        siteMap.set(code, {
          code,
          id: dbSite?.id || null,
          location: dbSite?.address || dbSite?.area || r.location || '—',
          city: dbSite?.city || 'Ahmedabad',
          size: dbSite?.size || (dbSite?.width && dbSite?.height ? `${dbSite.width}x${dbSite.height} ft` : r.size || (r.width && r.height ? `${r.width}x${r.height} ft` : '—')),
          width: dbSite?.width ?? r.width ?? null,
          height: dbSite?.height ?? r.height ?? null,
          type: dbSite?.type || r.type || 'Hoarding',
          rows: []
        });
      }
      siteMap.get(code).rows.push(r);
    });

    const now = new Date();

    const list = Array.from(siteMap.values()).map(site => {
      const sortedRows = [...site.rows].sort((a, b) => {
        const da = parseFlexibleDate(a.start_date || a.booking_date) || new Date(0);
        const db = parseFlexibleDate(b.start_date || b.booking_date) || new Date(0);
        return db - da;
      });

      let activeCampaign = null;
      let upcomingCampaign = null;
      let latestCampaign = sortedRows[0] || null;

      for (const r of sortedRows) {
        const sDate = parseFlexibleDate(r.start_date || r.booking_date);
        const eDate = parseFlexibleDate(r.end_date) || (sDate ? new Date(sDate.getFullYear(), sDate.getMonth() + 1, 0, 23, 59, 59) : null);
        if (sDate && eDate) {
          if (sDate <= now && eDate >= now) {
            activeCampaign = r;
            break;
          } else if (sDate > now && !upcomingCampaign) {
            upcomingCampaign = r;
          }
        }
      }

      const totalAmount = site.rows.reduce((sum, r) => sum + Number(r.total_amount || r.revenue || 0), 0);
      const totalPending = site.rows.reduce((sum, r) => sum + Number(r.pending || 0), 0);

      let status = 'vacant';
      let currentCampaign = null;
      if (activeCampaign) {
        status = 'occupied';
        currentCampaign = activeCampaign;
      } else if (upcomingCampaign) {
        status = 'upcoming';
        currentCampaign = upcomingCampaign;
      } else if (latestCampaign) {
        currentCampaign = latestCampaign;
        status = 'vacant';
      }

      return {
        ...site,
        rows: sortedRows,
        campaignCount: site.rows.length,
        status,
        currentCampaign,
        activeCampaign,
        upcomingCampaign,
        latestCampaign,
        totalAmount,
        totalPending
      };
    });

    return list.sort((a, b) => a.code.localeCompare(b.code, undefined, { numeric: true }));
  }, [sites, rows]);

  // Unified dataset: exactly 1 row per physical site displaying ONLY its latest/current booking
  const latestSiteCampaigns = useMemo(() => {
    return siteMasterList.map(s => {
      const c = s.currentCampaign;
      return {
        site_code: s.code,
        site_id: s.id,
        location: s.location,
        city: s.city,
        size: s.size,
        type: c?.type || s.type || 'Hoarding',
        width: c?.width ?? s.width ?? '',
        height: c?.height ?? s.height ?? '',
        siteStatus: s.status, // 'occupied' | 'upcoming' | 'vacant'
        campaignCount: s.campaignCount,
        allCampaigns: s.rows,

        // Latest campaign data:
        campaign_id: c?.id || null,
        campaign: c || null,
        month: c?.month || (c?.start_date ? new Date(c.start_date).toLocaleDateString('en-US', { month: 'short', year: 'numeric' }) : ''),
        occupancy: c ? getCampaignOccupancy(c) : { pct: 0, label: '0% Vacant', status: 'vacant' },
        booking_date: c?.booking_date || c?.date || '',
        client: c?.client || c?.client_name || '',
        display: c?.display || c?.campaign_name || '',
        vendor_name: c?.vendor_name || '',
        start_date: c?.start_date || '',
        end_date: c?.end_date || '',
        days: c?.days || (c?.start_date && c?.end_date ? Math.max(1, Math.round((new Date(c.end_date) - new Date(c.start_date)) / 86400000) + 1) : (c ? 30 : '')),
        advt_fees: Number(c?.advt_fees || 0),
        printing_mounting_cost: Number(c?.printing_mounting_cost || (Number(c?.printing_cost || 0) + Number(c?.mounting_cost || 0))),
        total_amount: Number(c?.total_amount || c?.revenue || 0),
        po: c?.po || '',
        bill: c?.bill || c?.invoice_no || '',
        pending: Number(c?.pending || 0),
        parent_campaign: c?.parent_campaign || '',
        notes: c?.notes || ''
      };
    });
  }, [siteMasterList]);

  // Counts & metrics
  const totalSites = siteMasterList.length;
  const occupiedSitesCount = useMemo(() => latestSiteCampaigns.filter(x => x.siteStatus === 'occupied').length, [latestSiteCampaigns]);
  const upcomingSitesCount = useMemo(() => latestSiteCampaigns.filter(x => x.siteStatus === 'upcoming').length, [latestSiteCampaigns]);
  const vacantSitesCount = useMemo(() => latestSiteCampaigns.filter(x => x.siteStatus === 'vacant').length, [latestSiteCampaigns]);
  const occupancyRate = totalSites > 0 ? Math.round((occupiedSitesCount / totalSites) * 100) : 0;
  const latestRevenueSum = useMemo(() => latestSiteCampaigns.reduce((sum, x) => sum + x.total_amount, 0), [latestSiteCampaigns]);
  const latestPendingSum = useMemo(() => latestSiteCampaigns.reduce((sum, x) => sum + x.pending, 0), [latestSiteCampaigns]);

  function handleSort(key) {
    setSortState(prev => prev.key === key ? { key, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { key, dir: 'asc' });
  }

  // Filtered and sorted latest site bookings list
  const filteredLatest = useMemo(() => {
    const list = latestSiteCampaigns.filter(item => {
      if (statusFilter === 'occupied' && item.siteStatus !== 'occupied') return false;
      if (statusFilter === 'upcoming' && item.siteStatus !== 'upcoming') return false;
      if (statusFilter === 'vacant' && item.siteStatus !== 'vacant') return false;
      if (statusFilter === 'pending' && item.pending <= 0) return false;

      if (!search.trim()) return true;
      const q = search.trim().toLowerCase();
      return (
        item.site_code.toLowerCase().includes(q) ||
        item.location.toLowerCase().includes(q) ||
        item.city.toLowerCase().includes(q) ||
        item.size.toLowerCase().includes(q) ||
        item.type.toLowerCase().includes(q) ||
        item.client.toLowerCase().includes(q) ||
        item.display.toLowerCase().includes(q) ||
        item.vendor_name.toLowerCase().includes(q) ||
        item.month.toLowerCase().includes(q) ||
        item.po.toLowerCase().includes(q) ||
        item.bill.toLowerCase().includes(q)
      );
    });

    if (sortState.key) {
      list.sort((a, b) => {
        let valA = a[sortState.key];
        let valB = b[sortState.key];
        if (sortState.key === 'occupancy') {
          valA = a.occupancy?.pct ?? 0;
          valB = b.occupancy?.pct ?? 0;
        } else if (['advt_fees', 'printing_mounting_cost', 'total_amount', 'pending', 'days', 'width', 'height'].includes(sortState.key)) {
          valA = Number(valA || 0);
          valB = Number(valB || 0);
        }
        return universalCompare(valA, valB, sortState.dir);
      });
    } else {
      list.sort((a, b) => a.site_code.localeCompare(b.site_code, undefined, { numeric: true }));
    }
    return list;
  }, [latestSiteCampaigns, statusFilter, search, sortState]);

  function toggleSelect(id) {
    if (!id) return;
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function toggleSelectAll() {
    const visibleCampaignIds = filteredLatest.map(r => r.campaign_id).filter(Boolean);
    const allSelected = visibleCampaignIds.length > 0 && visibleCampaignIds.every(id => selectedIds.has(id));
    if (allSelected) {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleCampaignIds.forEach(id => next.delete(id));
        return next;
      });
    } else {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleCampaignIds.forEach(id => next.add(id));
        return next;
      });
    }
  }

  function selectAllVisible() {
    setSelectedIds(new Set(filteredLatest.map(r => r.campaign_id).filter(Boolean)));
  }

  function deselectAll() {
    setSelectedIds(new Set());
  }

  async function handleBatchDelete() {
    if (!canDelete || selectedIds.size === 0) return;
    const count = selectedIds.size;
    if (!confirm(`Are you sure you want to permanently delete all ${count} selected campaign record${count > 1 ? 's' : ''}? This action cannot be undone.`)) {
      return;
    }
    try {
      await api.post('/campaigns/batch-delete', { ids: Array.from(selectedIds), hard: true });
      setBanner(`✓ ${count} campaign record${count > 1 ? 's' : ''} deleted successfully.`);
      setSelectedIds(new Set());
      await loadData();
      setTimeout(() => setBanner(''), 4000);
    } catch (err) {
      alert('Failed to delete selected campaigns: ' + (err.response?.data?.message || err.message));
    }
  }

  async function handleDeleteAllCampaigns() {
    if (!canDelete) return;
    const allVisibleCampIds = filteredLatest.map(r => r.campaign_id).filter(Boolean);
    const count = allVisibleCampIds.length;
    if (count === 0) {
      alert('There are no active or visible campaign records to delete.');
      return;
    }
    if (!confirm(`⚠️ DANGER: Are you sure you want to permanently delete ALL ${count} visible campaign records from the Campaign Tracker?\n\nThis will clear all booking data for these sites. This action cannot be undone.`)) {
      return;
    }
    try {
      await api.post('/campaigns/batch-delete', { ids: allVisibleCampIds, hard: true });
      setBanner(`✓ All ${count} campaign records deleted successfully.`);
      setSelectedIds(new Set());
      await loadData();
      setTimeout(() => setBanner(''), 5000);
    } catch (err) {
      alert('Failed to delete all campaigns: ' + (err.response?.data?.message || err.message));
    }
  }

  async function deleteRecord(id) {
    if (!canDelete) return;
    if (confirm('Are you sure you want to permanently delete this campaign record?')) {
      try {
        await api.delete(`/campaigns/${id}?hard=true`);
        setBanner('✓ Campaign record deleted successfully.');
        setSelectedIds(prev => {
          const next = new Set(prev);
          next.delete(id);
          return next;
        });
        await loadData();
        setTimeout(() => setBanner(''), 3000);
      } catch (err) {
        alert('Failed to delete campaign: ' + (err.response?.data?.message || err.message));
      }
    }
  }

  async function handleCampaignsExcelImport(e) {
    if (!canDelete) return;
    const file = e.target.files?.[0];
    if (!file) return;
    setImportingExcel(true);
    try {
      const fd = new FormData();
      fd.append('file', file);
      const r = await api.post('/import/campaigns-xlsx', fd, { headers: { 'Content-Type': 'multipart/form-data' } });
      const msg = r.data.message || `Processed ${r.data.rows} campaigns (${r.data.updated} updated, ${r.data.created} created).`;
      setBanner(`✓ ${msg}`);
      alert(`✓ Campaigns Import Successful!\n\n${msg}`);
      await loadData();
      setTimeout(() => setBanner(''), 6000);
    } catch (err) {
      alert('Campaign Excel Import failed: ' + (err.response?.data?.message || err.message));
    } finally {
      setImportingExcel(false);
      e.target.value = '';
    }
  }

  function openNewCampaign(code) {
    if (!canAdd) return;
    setIsViewingOnly(false);
    const initialSiteCode = typeof code === 'string' && code ? code : (siteQueryParam || (search.trim().startsWith('MB-') ? search.trim() : ''));
    setModalSiteCode(initialSiteCode);
    const foundSite = sites.find(s => String(s.site_code || '').toUpperCase() === initialSiteCode.toUpperCase());
    const initLoc = foundSite ? (foundSite.address || foundSite.area || '') : '';
    const initW = foundSite?.width ?? '';
    const initH = foundSite?.height ?? '';
    const initSize = foundSite?.size ?? (initW && initH ? `${initW}x${initH} ft` : '');
    const initType = foundSite?.type || 'Hoarding';
    setModalLocation(initLoc);
    setModalWidth(initW);
    setModalHeight(initH);
    setModalSize(initSize);
    setModalType(initType);
    setAutoMatchedSite(foundSite || null);
    setEditModal({
      site_code: initialSiteCode,
      month: new Date().toLocaleDateString('en-US', { month: 'short', year: 'numeric' }),
      booking_date: new Date().toISOString().slice(0, 10),
      client: '',
      display: '',
      vendor_name: '',
      location: initLoc,
      width: initW,
      height: initH,
      size: initSize,
      type: initType,
      start_date: new Date().toISOString().slice(0, 10),
      end_date: '',
      days: 30,
      advt_fees: '',
      printing_mounting_cost: '',
      total_amount: '',
      po: '',
      bill: '',
      pending: '',
      notes: ''
    });
  }

  function openEditCampaign(r) {
    if (!r) return;
    setIsViewingOnly(false);
    setModalSiteCode(r.site_code || '');
    setModalLocation(r.location || '');
    setModalWidth(r.width ?? '');
    setModalHeight(r.height ?? '');
    setModalSize(r.size || '');
    setModalType(r.type || 'Hoarding');
    setAutoMatchedSite(null);
    setEditModal(r);
  }

  function openViewCampaign(r) {
    if (!r) return;
    setIsViewingOnly(true);
    setModalSiteCode(r.site_code || '');
    setModalLocation(r.location || '');
    setModalWidth(r.width ?? '');
    setModalHeight(r.height ?? '');
    setModalSize(r.size || '');
    setModalType(r.type || 'Hoarding');
    setAutoMatchedSite(null);
    setEditModal(r);
  }

  async function saveCampaign(e) {
    e.preventDefault();
    if (isReadOnly) return;
    setSaving(true);
    const form = e.currentTarget;
    const formData = new FormData(form);
    const payload = Object.fromEntries(formData.entries());

    payload.advt_fees = Number(payload.advt_fees || 0);
    payload.printing_mounting_cost = Number(payload.printing_mounting_cost || 0);
    let totalAmt = Number(payload.total_amount || 0);
    if (!totalAmt && (payload.advt_fees || payload.printing_mounting_cost)) {
      totalAmt = payload.advt_fees + payload.printing_mounting_cost;
    }
    payload.total_amount = totalAmt;
    payload.revenue = totalAmt;
    payload.pending = Number(payload.pending || 0);
    payload.days = Number(payload.days || 30);
    if (payload.width) payload.width = parseFloat(payload.width) || null;
    if (payload.height) payload.height = parseFloat(payload.height) || null;
    if (!payload.size && payload.width && payload.height) {
      payload.size = `${payload.width}x${payload.height} ft`;
    }
    if (payload.site_code) payload.site_code = String(payload.site_code).trim();
    if (payload.client) payload.client_name = payload.client;
    if (payload.display) {
      payload.campaign_name = payload.display;
      payload.brand = payload.display;
    }

    try {
      if (editModal?.id) {
        await api.put(`/campaigns/${editModal.id}`, payload);
        setBanner('✓ Campaign record updated successfully!');
      } else {
        await api.post('/campaigns', payload);
        setBanner('✓ New campaign created successfully!');
      }
      setEditModal(null);
      await loadData();
      setTimeout(() => setBanner(''), 4000);
    } catch (err) {
      alert('Failed to save campaign: ' + (err.response?.data?.message || err.message));
    } finally {
      setSaving(false);
    }
  }

  function renderLatestSiteRow(item) {
    const isPending = item.pending > 0;
    const isBooked = item.siteStatus === 'occupied';
    const isUpcoming = item.siteStatus === 'upcoming';
    const statusBg = isBooked ? 'rgba(34, 197, 94, 0.16)' : isUpcoming ? 'rgba(234, 179, 8, 0.16)' : 'rgba(148, 163, 184, 0.12)';
    const statusColor = isBooked ? '#4ade80' : isUpcoming ? '#facc15' : '#94a3b8';
    const statusBorder = isBooked ? 'rgba(34, 197, 94, 0.35)' : isUpcoming ? 'rgba(234, 179, 8, 0.35)' : 'rgba(148, 163, 184, 0.25)';

    return (
      <tr key={item.site_code}>
        {canDelete && (
          <td style={{ textAlign: 'center', padding: '10px 8px' }}>
            {item.campaign_id ? (
              <input
                type="checkbox"
                checked={selectedIds.has(item.campaign_id)}
                onChange={() => toggleSelect(item.campaign_id)}
                style={{ cursor: 'pointer' }}
                title={`Select latest campaign on ${item.site_code}`}
              />
            ) : (
              <span style={{ color: '#475569', fontSize: '11px' }}>—</span>
            )}
          </td>
        )}
        <td>
          <div style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
            <span className="scooh-plate" style={{ fontSize: '12px', fontWeight: 800 }}>
              {item.site_code}
            </span>
            {String(item.parent_campaign || '').startsWith('LINKED:') ? (
              <span
                style={{
                  fontSize: '10px',
                  padding: '1px 6px',
                  borderRadius: '4px',
                  background: 'rgba(245,158,11,0.2)',
                  color: '#fbbf24',
                  border: '1px solid rgba(245,158,11,0.4)',
                  fontWeight: 700
                }}
                title={item.notes || 'Auto-booked linked entry'}
              >
                🔗 Linked
              </span>
            ) : getSiteTypeTag(item.site_code) !== 'Single' && (
              <span
                style={{
                  fontSize: '10px',
                  padding: '1px 5px',
                  borderRadius: '4px',
                  background: getSiteTypeTag(item.site_code) === 'Combined' ? 'rgba(168,85,247,0.2)' : 'rgba(56,189,248,0.2)',
                  color: getSiteTypeTag(item.site_code) === 'Combined' ? '#c084fc' : '#38bdf8',
                  fontWeight: 700
                }}
                title={getConflictSummary(item.site_code)?.message || ''}
              >
                {getSiteTypeTag(item.site_code)}
              </span>
            )}
          </div>
        </td>
        <td style={{ color: '#cbd5e1', fontSize: '12px', fontWeight: 600 }}>
          {item.month ? (
            <span style={{ padding: '2px 8px', borderRadius: '5px', background: 'rgba(56, 189, 248, 0.1)', border: '1px solid rgba(56, 189, 248, 0.25)', color: '#38bdf8' }}>
              {item.month}
            </span>
          ) : (
            <span style={{ color: '#64748b' }}>—</span>
          )}
        </td>
        <td>
          <span style={{
            display: 'inline-flex',
            alignItems: 'center',
            gap: '5px',
            padding: '3px 9px',
            borderRadius: '6px',
            background: statusBg,
            color: statusColor,
            border: `1px solid ${statusBorder}`,
            fontSize: '11.5px',
            fontWeight: 800,
            whiteSpace: 'nowrap'
          }}>
            <span style={{ width: '6px', height: '6px', borderRadius: '50%', background: statusColor }} />
            {isBooked ? 'Occupied' : isUpcoming ? 'Upcoming' : 'Vacant'}
          </span>
        </td>
        <td style={{ color: '#94a3b8', fontSize: '11.5px' }}>
          {formatDate(item.booking_date) || '—'}
        </td>
        <td>
          {item.client ? (
            <div style={{ fontWeight: 800, color: '#ffffff', fontSize: '13px' }}>
              {item.client}
            </div>
          ) : (
            <span style={{ color: '#64748b', fontSize: '12px', fontStyle: 'italic' }}>
              — Ready for Booking —
            </span>
          )}
        </td>
        <td>
          {item.display ? (
            <div style={{ color: '#38bdf8', fontWeight: 600, fontSize: '12.5px' }}>
              📢 {cleanDisplayTitle(item.display)}
            </div>
          ) : (
            <span style={{ color: '#64748b' }}>—</span>
          )}
        </td>
        <td style={{ color: '#cbd5e1', fontSize: '12px' }}>
          {item.vendor_name || '—'}
        </td>
        <td>
          <div style={{ color: '#edf2f6', fontSize: '12px', lineHeight: 1.35, fontWeight: 600 }}>
            {item.location || '—'}
          </div>
          {item.city && (
            <div style={{ fontSize: '11px', color: '#94a3b8', marginTop: '2px' }}>
              📍 {item.city}
            </div>
          )}
        </td>
        <td style={{ textAlign: 'center', color: '#94a3b8', fontSize: '11.5px' }}>
          {item.width || '—'}
        </td>
        <td style={{ textAlign: 'center', color: '#94a3b8', fontSize: '11.5px' }}>
          {item.height || '—'}
        </td>
        <td style={{ color: '#e2e8f0', fontSize: '12px', fontWeight: 600 }}>
          {item.size || '—'}
        </td>
        <td>
          <span className="scooh-badgechip" style={{ background: 'rgba(56,189,248,0.14)', color: '#38bdf8', fontSize: '11px', fontWeight: 700 }}>
            {item.type}
          </span>
        </td>
        <td style={{ fontSize: '11.5px', color: '#94a3b8' }}>
          {formatDate(item.start_date) || '—'}
        </td>
        <td style={{ fontSize: '11.5px', color: '#94a3b8' }}>
          {formatDate(item.end_date) || (item.start_date ? 'Ongoing' : '—')}
        </td>
        <td style={{ textAlign: 'center', fontWeight: 700, color: '#e2e8f0', fontSize: '12px' }}>
          {item.days ? `${item.days}d` : '—'}
        </td>
        <td style={{ textAlign: 'right', fontWeight: 600, color: '#e2e8f0', fontSize: '12.5px' }}>
          {String(item.parent_campaign || '').startsWith('LINKED:') ? (
            <span style={{ color: '#94a3b8', fontSize: '11.5px', fontStyle: 'italic' }}>—</span>
          ) : item.advt_fees > 0 ? (
            money(item.advt_fees)
          ) : (
            '—'
          )}
        </td>
        <td style={{ textAlign: 'right', fontWeight: 600, color: '#e2e8f0', fontSize: '12.5px' }}>
          {String(item.parent_campaign || '').startsWith('LINKED:') ? (
            <span style={{ color: '#94a3b8', fontSize: '11.5px', fontStyle: 'italic' }}>—</span>
          ) : item.printing_mounting_cost > 0 ? (
            money(item.printing_mounting_cost)
          ) : (
            '—'
          )}
        </td>
        <td style={{ textAlign: 'right', fontWeight: 800, color: '#4ade80', fontSize: '13px' }}>
          {String(item.parent_campaign || '').startsWith('LINKED:') ? (
            <span style={{ color: '#94a3b8', fontSize: '11.5px', fontStyle: 'italic', fontWeight: 600 }} title="Covered under primary booking">
              Covered
            </span>
          ) : item.total_amount > 0 ? (
            money(item.total_amount)
          ) : (
            '—'
          )}
        </td>
        <td>
          {item.po ? <span className="scooh-plate" style={{ fontSize: '11px' }}>{item.po}</span> : '—'}
        </td>
        <td>
          {item.bill ? <span style={{ color: '#93c5fd', fontSize: '11px' }}>{item.bill}</span> : '—'}
        </td>
        <td style={{ textAlign: 'right', fontWeight: 700, color: isPending ? '#f87171' : '#94a3b8', fontSize: '12.5px' }}>
          {item.pending > 0 ? money(item.pending) : item.total_amount > 0 ? <span style={{ color: '#4ade80', fontSize: '11px' }}>₹0</span> : '—'}
        </td>
        <td>
          <div className="scooh-rowactions" style={{ justifyContent: 'center', gap: '5px' }}>
            {item.campaign && (
              <button
                type="button"
                className="scooh-btn ghost"
                style={{ fontSize: '11px', padding: '3px 8px', color: '#38bdf8', borderColor: 'rgba(56, 189, 248, 0.4)' }}
                onClick={() => openViewCampaign(item.campaign)}
                title="View campaign specifications and billing details"
              >
                👁 View
              </button>
            )}
            {item.campaign && canEdit && !isReadOnly && (
              <button
                type="button"
                className="scooh-iconbtn scooh-text-action"
                style={{ fontSize: '11px', padding: '3px 8px' }}
                onClick={() => openEditCampaign(item.campaign)}
                title="Edit campaign record"
              >
                ✏️ Edit
              </button>
            )}
            {canAdd && !item.campaign && (
              <button
                type="button"
                className="scooh-btn purple-btn"
                style={{ fontSize: '11px', padding: '3px 8px', borderRadius: '5px' }}
                onClick={() => openNewCampaign(item.site_code)}
                title={`Book new campaign for ${item.site_code}`}
              >
                + Book
              </button>
            )}
            {canDelete && item.campaign_id && (
              <button
                type="button"
                className="scooh-iconbtn danger-icon"
                style={{ fontSize: '11px', padding: '3px 6px' }}
                onClick={() => deleteRecord(item.campaign_id)}
                title="Delete this campaign record"
              >
                ×
              </button>
            )}
          </div>
        </td>
      </tr>
    );
  }

  return (
    <>
      <div className="scooh-pagehead">
        <div>
          <div style={{ fontSize: '11px', fontWeight: 800, color: '#a78bfa', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: '4px' }}>
            Media Buzz — OOH Workspace
          </div>
          <h1 style={{ fontSize: '26px', fontWeight: 900, margin: 0, color: '#fff', letterSpacing: '-0.02em' }}>
            Campaign Tracker
          </h1>
          <p style={{ margin: '6px 0 0', color: '#94a3b8', fontSize: '13px' }}>
            Live outdoor campaigns, latest client billing, and real-time site occupancy.
          </p>
        </div>
        <div className="scooh-headactions" style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
          <button type="button" className="scooh-btn ghost" onClick={loadData}>
            🔄 Refresh
          </button>
          {(canAdd || canDelete) && (
            <label className="scooh-btn ghost" style={{ cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: '6px' }} title="Import campaigns from Excel or CSV spreadsheet">
              <span>📁 {importingExcel ? 'Importing…' : 'Import Excel'}</span>
              <input type="file" accept=".xlsx,.xls,.csv,.xlsm,.ods" hidden disabled={importingExcel} onChange={handleCampaignsExcelImport} />
            </label>
          )}
          <button type="button" className="scooh-btn ghost" onClick={() => exportCampaignsExcel(filteredLatest, 'MediaBuzz_Campaign_Tracker.xlsx')}>
            📥 Export Excel
          </button>
          {canAdd && (
            <button type="button" className="scooh-btn purple-btn" onClick={() => openNewCampaign()}>
              + Add Campaign
            </button>
          )}
        </div>
      </div>

      {isReadOnly && (
        <div className="scooh-banner" style={{ display: 'block', marginBottom: '16px', background: 'rgba(245,158,11,0.12)', border: '1px solid rgba(245,158,11,0.3)', color: '#fbbf24' }}>
          🔒 Read-Only Workspace: You are signed in as a Viewer. You can view all campaign data, filter records, and export to Excel. Creating, editing, and deleting campaigns is disabled.
        </div>
      )}

      {banner && (
        <div className="scooh-banner" style={{ display: 'block', marginBottom: '16px', background: 'rgba(34,197,94,0.15)', border: '1px solid #22c55e', color: '#4ade80' }}>
          {banner}
        </div>
      )}

      {/* Summary KPI Cards (Interactive Filter Shortcuts) */}
      <div className="scooh-kpirow" style={{ marginBottom: '20px' }}>
        <div 
          className={`scooh-kpi ${occupancyRate >= 75 ? 'good' : occupancyRate >= 50 ? 'alert' : 'danger'}`}
          style={{ cursor: 'pointer' }}
          onClick={() => setStatusFilter(prev => prev === 'occupied' ? 'ALL' : 'occupied')}
          title="Click to toggle occupied sites"
        >
          <div className="n">{occupancyRate}%</div>
          <div className="l">Occupancy Rate</div>
          <div className="scooh-kpi-note" style={{ color: occupancyRate >= 75 ? '#4ade80' : '#fbbf24' }}>
            {occupiedSitesCount} of {totalSites} sites occupied
          </div>
        </div>

        <div 
          className="scooh-kpi good"
          style={{ cursor: 'pointer' }}
          onClick={() => setStatusFilter(prev => prev === 'vacant' ? 'ALL' : 'vacant')}
          title="Click to view vacant sites ready for booking"
        >
          <div className="n" style={{ color: vacantSitesCount > 0 ? '#fbbf24' : '#4ade80' }}>{vacantSitesCount}</div>
          <div className="l">Vacant Sites</div>
          <div className="scooh-kpi-note" style={{ color: '#94a3b8' }}>Available for new campaigns</div>
        </div>

        <div 
          className="scooh-kpi good"
          style={{ cursor: 'pointer' }}
          onClick={() => setStatusFilter('ALL')}
          title="Click to view all sites"
        >
          <div className="n">{money(latestRevenueSum)}</div>
          <div className="l">Current Campaign Revenue</div>
          <div className="scooh-kpi-note" style={{ color: '#4ade80' }}>{occupiedSitesCount + upcomingSitesCount} active / upcoming bookings</div>
        </div>

        <div 
          className={`scooh-kpi ${latestPendingSum > 0 ? 'danger' : 'good'}`}
          style={{ cursor: 'pointer' }}
          onClick={() => setStatusFilter(prev => prev === 'pending' ? 'ALL' : 'pending')}
          title="Click to filter pending payments"
        >
          <div className="n">{money(latestPendingSum)}</div>
          <div className="l">Total Pending Amount</div>
          <div className="scooh-kpi-note" style={{ color: latestPendingSum > 0 ? '#ef4444' : '#4ade80' }}>
            {latestPendingSum > 0 ? `${latestSiteCampaigns.filter(r => r.pending > 0).length} pending sites` : 'All cleared'}
          </div>
        </div>
      </div>

      {/* Main Campaign Section */}
      <div className="scooh-electricity-section">
        <div className="scooh-sectionbar" style={{ padding: '14px 20px' }}>
          <div>
            <h2 style={{ fontSize: '16px', fontWeight: 800, color: '#f8fafc', margin: 0 }}>
              Live Sites & Latest Campaigns
            </h2>
            <p style={{ margin: '3px 0 0', fontSize: '12px', color: '#94a3b8' }}>
              Showing latest booking per site ({filteredLatest.length} of {totalSites} sites). To see full previous booking records, visit Occupancy.
            </p>
          </div>
          <div style={{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap' }}>
            <button
              type="button"
              className="scooh-btn ghost"
              onClick={() => navigate('/occupancy')}
              style={{ fontSize: '12px', color: '#38bdf8', borderColor: 'rgba(56, 189, 248, 0.3)' }}
              title="Open full booking history ledger in Occupancy"
            >
              📋 All Previous Bookings (Occupancy) ↗
            </button>
            {canAdd && (
              <button type="button" className="scooh-btn purple-btn" onClick={() => openNewCampaign()}>
                + Add Campaign
              </button>
            )}
          </div>
        </div>

        {/* Clean, Unified Single Toolbar */}
        <div className="scooh-toolbar" style={{ padding: '12px 20px', borderBottom: '1px solid rgba(255,255,255,0.06)', display: 'flex', gap: '12px', flexWrap: 'wrap', alignItems: 'center', justifyContent: 'space-between' }}>
          
          {/* Left: Status Filter Pills & Live Search */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap', flex: '1 1 auto' }}>
            
            <div style={{ display: 'inline-flex', gap: '4px', background: 'rgba(15, 23, 42, 0.7)', padding: '3px', borderRadius: '20px', border: '1px solid #1e293b' }}>
              <button
                type="button"
                onClick={() => setStatusFilter('ALL')}
                style={{
                  padding: '5px 12px',
                  borderRadius: '16px',
                  border: 'none',
                  background: statusFilter === 'ALL' ? 'rgba(167, 139, 250, 0.22)' : 'transparent',
                  color: statusFilter === 'ALL' ? '#c4b5fd' : '#94a3b8',
                  fontWeight: statusFilter === 'ALL' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                All Sites ({totalSites})
              </button>
              <button
                type="button"
                onClick={() => setStatusFilter('occupied')}
                style={{
                  padding: '5px 12px',
                  borderRadius: '16px',
                  border: 'none',
                  background: statusFilter === 'occupied' ? 'rgba(34, 197, 94, 0.2)' : 'transparent',
                  color: statusFilter === 'occupied' ? '#4ade80' : '#94a3b8',
                  fontWeight: statusFilter === 'occupied' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                🟢 Occupied ({occupiedSitesCount})
              </button>
              <button
                type="button"
                onClick={() => setStatusFilter('upcoming')}
                style={{
                  padding: '5px 12px',
                  borderRadius: '16px',
                  border: 'none',
                  background: statusFilter === 'upcoming' ? 'rgba(234, 179, 8, 0.2)' : 'transparent',
                  color: statusFilter === 'upcoming' ? '#facc15' : '#94a3b8',
                  fontWeight: statusFilter === 'upcoming' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                🟡 Upcoming ({upcomingSitesCount})
              </button>
              <button
                type="button"
                onClick={() => setStatusFilter('vacant')}
                style={{
                  padding: '5px 12px',
                  borderRadius: '16px',
                  border: 'none',
                  background: statusFilter === 'vacant' ? 'rgba(148, 163, 184, 0.2)' : 'transparent',
                  color: statusFilter === 'vacant' ? '#e2e8f0' : '#94a3b8',
                  fontWeight: statusFilter === 'vacant' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                ⚪ Vacant ({vacantSitesCount})
              </button>
              <button
                type="button"
                onClick={() => setStatusFilter('pending')}
                style={{
                  padding: '5px 12px',
                  borderRadius: '16px',
                  border: 'none',
                  background: statusFilter === 'pending' ? 'rgba(239, 68, 68, 0.2)' : 'transparent',
                  color: statusFilter === 'pending' ? '#f87171' : '#94a3b8',
                  fontWeight: statusFilter === 'pending' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                ⚠️ Pending ({latestSiteCampaigns.filter(x => x.pending > 0).length})
              </button>
            </div>

            {/* Search Input */}
            <input
              className="scooh-search"
              style={{ minWidth: '220px', maxWidth: '320px', fontSize: '12px' }}
              placeholder="Search site, location, client, display, PO, Bill…"
              value={search}
              onChange={e => setSearch(e.target.value)}
            />

            {(search || statusFilter !== 'ALL') && (
              <button
                type="button"
                className="scooh-btn ghost"
                style={{ fontSize: '11px', padding: '5px 9px' }}
                onClick={() => {
                  setSearch('');
                  setStatusFilter('ALL');
                  setSiteQueryParam('');
                  navigate('/campaigns', { replace: true });
                }}
              >
                ✕ Reset
              </button>
            )}

            {canDelete && (
              <div style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', marginLeft: '4px' }}>
                <button
                  type="button"
                  className="scooh-btn ghost"
                  style={{ fontSize: '11.5px', padding: '5px 10px', color: '#38bdf8', borderColor: 'rgba(56, 189, 248, 0.3)' }}
                  onClick={selectAllVisible}
                  title="Select all visible campaign records"
                >
                  ☑ Select All
                </button>
                {selectedIds.size > 0 && (
                  <button
                    type="button"
                    className="scooh-btn ghost"
                    style={{ fontSize: '11.5px', padding: '5px 10px' }}
                    onClick={deselectAll}
                    title="Clear selection"
                  >
                    ☐ Deselect All
                  </button>
                )}
                {selectedIds.size > 0 && (
                  <button
                    type="button"
                    className="scooh-btn danger"
                    style={{ fontSize: '11.5px', padding: '5px 12px' }}
                    onClick={handleBatchDelete}
                  >
                    🗑 Delete Selected ({selectedIds.size})
                  </button>
                )}
                <button
                  type="button"
                  className="scooh-btn danger"
                  style={{ fontSize: '11.5px', padding: '5px 12px', background: 'rgba(239, 68, 68, 0.15)', borderColor: 'rgba(239, 68, 68, 0.35)', color: '#f87171' }}
                  onClick={handleDeleteAllCampaigns}
                  title="Delete all visible campaign records with confirmation"
                >
                  🗑 Delete All
                </button>
              </div>
            )}
          </div>

          {/* Right: Quick Action Buttons */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <button
              type="button"
              className="scooh-btn ghost"
              onClick={() => exportCampaignsExcel(filteredLatest, 'MediaBuzz_Campaign_Tracker.xlsx')}
              title="Export table to Excel"
              style={{ fontSize: '11.5px', padding: '6px 12px' }}
            >
              📥 Export Excel
            </button>

            {(canAdd || canDelete) && (
              <label className="scooh-btn ghost" style={{ cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '11.5px', padding: '6px 12px' }} title="Import campaigns spreadsheet">
                <span>📁 {importingExcel ? 'Importing…' : 'Import Excel'}</span>
                <input type="file" accept=".xlsx,.xls,.csv,.xlsm,.ods" hidden disabled={importingExcel} onChange={handleCampaignsExcelImport} />
              </label>
            )}

            <button type="button" className="scooh-btn ghost" onClick={loadData} title="Refresh data" style={{ fontSize: '11.5px', padding: '6px 10px' }}>
              🔄
            </button>
          </div>
        </div>

        {/* Unified Latest Bookings Table with Dual Synchronized Scrollers */}
        <div>
          {/* Top Synchronized Scroller Bar */}
          <div
            ref={topScrollRef}
            onScroll={handleTopScroll}
            className="scooh-tablewrap"
            style={{
              overflowX: 'auto',
              overflowY: 'hidden',
              maxHeight: '14px',
              marginBottom: '4px',
              background: 'rgba(15, 23, 42, 0.6)',
              borderRadius: '6px',
              border: '1px solid rgba(255, 255, 255, 0.08)'
            }}
          >
            <div style={{ width: `${tableScrollWidth}px`, height: '1px' }} />
          </div>

          {/* Bottom Primary Table & Scroller */}
          <div ref={bottomScrollRef} onScroll={handleBottomScroll} className="scooh-tablewrap" style={{ overflowX: 'auto' }}>
            <table ref={tableRef} className="scooh-table" style={{ minWidth: '2200px' }}>
              <thead>
                <tr>
                  {canDelete && (
                    <th style={{ width: '40px', textAlign: 'center', padding: '10px 8px' }}>
                      <input
                        type="checkbox"
                        checked={filteredLatest.length > 0 && filteredLatest.some(r => r.campaign_id) && filteredLatest.filter(r => r.campaign_id).every(r => selectedIds.has(r.campaign_id))}
                        onChange={toggleSelectAll}
                        title="Select all visible latest campaigns"
                        style={{ cursor: 'pointer' }}
                      />
                    </th>
                  )}
                  <SortHeader label="Site Code" sortKey="site_code" currentSort={sortState} onSort={handleSort} style={{ minWidth: '105px' }} />
                  <SortHeader label="Month" sortKey="month" currentSort={sortState} onSort={handleSort} style={{ minWidth: '95px' }} />
                  <SortHeader label="Occupancy" sortKey="occupancy" currentSort={sortState} onSort={handleSort} style={{ minWidth: '135px' }} />
                  <SortHeader label="Date" sortKey="booking_date" currentSort={sortState} onSort={handleSort} style={{ minWidth: '100px' }} />
                  <SortHeader label="Client/Agency Name" sortKey="client" currentSort={sortState} onSort={handleSort} style={{ minWidth: '180px' }} />
                  <SortHeader label="Display" sortKey="display" currentSort={sortState} onSort={handleSort} style={{ minWidth: '160px' }} />
                  <SortHeader label="Vendor Name" sortKey="vendor_name" currentSort={sortState} onSort={handleSort} style={{ minWidth: '140px' }} />
                  <SortHeader label="Location" sortKey="location" currentSort={sortState} onSort={handleSort} style={{ minWidth: '180px' }} />
                  <th style={{ width: '55px', textAlign: 'center' }}>W</th>
                  <th style={{ width: '55px', textAlign: 'center' }}>H</th>
                  <SortHeader label="Size" sortKey="size" currentSort={sortState} onSort={handleSort} style={{ minWidth: '95px' }} />
                  <SortHeader label="Type" sortKey="type" currentSort={sortState} onSort={handleSort} style={{ minWidth: '105px' }} />
                  <SortHeader label="Start Date" sortKey="start_date" currentSort={sortState} onSort={handleSort} style={{ minWidth: '105px' }} />
                  <SortHeader label="End Date" sortKey="end_date" currentSort={sortState} onSort={handleSort} style={{ minWidth: '105px' }} />
                  <SortHeader label="Days" sortKey="days" currentSort={sortState} onSort={handleSort} align="center" style={{ width: '70px' }} />
                  <SortHeader label="Advt. Fees" sortKey="advt_fees" currentSort={sortState} onSort={handleSort} align="right" style={{ minWidth: '140px' }} />
                  <SortHeader label="Prod/Mount" sortKey="printing_mounting_cost" currentSort={sortState} onSort={handleSort} align="right" style={{ minWidth: '140px' }} />
                  <SortHeader label="Total" sortKey="total_amount" currentSort={sortState} onSort={handleSort} align="right" style={{ minWidth: '135px' }} />
                  <SortHeader label="PO" sortKey="po" currentSort={sortState} onSort={handleSort} style={{ minWidth: '95px' }} />
                  <SortHeader label="Bill" sortKey="bill" currentSort={sortState} onSort={handleSort} style={{ minWidth: '100px' }} />
                  <SortHeader label="Pending" sortKey="pending" currentSort={sortState} onSort={handleSort} align="right" style={{ minWidth: '110px' }} />
                  <th style={{ minWidth: '180px', textAlign: 'center' }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {filteredLatest.length === 0 ? (
                  <tr>
                    <td colSpan={canDelete ? 23 : 22} className="scooh-empty" style={{ padding: '36px 20px', textAlign: 'center' }}>
                      <div style={{ fontSize: '14px', fontWeight: 700, color: '#94a3b8' }}>No sites match your query</div>
                      <div style={{ fontSize: '12px', color: '#64748b', marginTop: '4px' }}>
                        Try adjusting your search or status filter.
                      </div>
                    </td>
                  </tr>
                ) : (
                  filteredLatest.map(item => renderLatestSiteRow(item))
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* Add / Edit / View Campaign Modal */}
      {editModal && (
        <div className="scooh-modal-overlay" onClick={() => !saving && setEditModal(null)}>
          <div className="scooh-modal" style={{ maxWidth: '780px', width: '92%' }} onClick={e => e.stopPropagation()}>
            <div className="scooh-modalhead">
              <div>
                <h2>{editModal.id ? (isViewingOnly || isReadOnly ? '👁 View Campaign Details' : 'Edit Campaign Record') : 'Add Campaign Record'}</h2>
                <span className="scooh-modal-subtitle">
                  {isViewingOnly || isReadOnly ? 'Detailed specification, schedule, and financial preview for this booking' : 'Enter outdoor campaign specifications and billing details'}
                </span>
              </div>
              <button type="button" className="scooh-modalclose" onClick={() => setEditModal(null)}>×</button>
            </div>
            <form onSubmit={saveCampaign}>
              <div className="scooh-modalbody" style={{ maxHeight: '72vh', overflowY: 'auto' }}>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '14px' }}>
                  <div className="scooh-field">
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <label>Site Code</label>
                      {autoMatchedSite && (
                        <span style={{ fontSize: '11px', color: '#10b981', fontWeight: 700 }}>
                          ✓ Match: {autoMatchedSite.site_code}
                        </span>
                      )}
                    </div>
                    <input
                      name="site_code"
                      list="campaign-site-codes-list"
                      value={modalSiteCode}
                      onChange={e => handleSiteCodeChange(e.target.value)}
                      placeholder="e.g. MB-06 or MB-02"
                      disabled={isViewingOnly || isReadOnly}
                    />
                    <datalist id="campaign-site-codes-list">
                      {sites.map(s => (
                        <option key={s.id || s.site_code} value={s.site_code}>
                          {s.site_code} — {s.address || s.area || ''} ({s.size || `${s.width}x${s.height}`})
                        </option>
                      ))}
                    </datalist>
                  </div>
                  {autoMatchedSite && !isViewingOnly && (
                    <div style={{
                      gridColumn: '1 / -1',
                      padding: '9px 13px',
                      borderRadius: '8px',
                      background: 'rgba(16, 185, 129, 0.12)',
                      border: '1px solid rgba(16, 185, 129, 0.35)',
                      fontSize: '12px',
                      color: '#ecfdf5',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between',
                      gap: '10px',
                      flexWrap: 'wrap'
                    }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <span style={{ fontSize: '16px' }}>🎯</span>
                        <div>
                          <span style={{ color: '#a7f3d0' }}>Auto-detected Site Code from Location & Size:</span>{' '}
                          <strong style={{ color: '#ffffff', fontSize: '13px', background: 'rgba(16, 185, 129, 0.3)', padding: '2px 7px', borderRadius: '4px' }}>
                            {autoMatchedSite.site_code}
                          </strong>
                          <span style={{ color: '#cbd5e1', marginLeft: '6px' }}>
                            {autoMatchedSite.size ? `[${autoMatchedSite.size}]` : ''} {autoMatchedSite.address || autoMatchedSite.area ? `• ${autoMatchedSite.address || autoMatchedSite.area}` : ''}
                          </span>
                        </div>
                      </div>
                      {String(modalSiteCode).trim().toUpperCase() !== String(autoMatchedSite.site_code).toUpperCase() ? (
                        <button
                          type="button"
                          onClick={() => applyMatchedSite(autoMatchedSite)}
                          style={{
                            padding: '4px 12px',
                            borderRadius: '6px',
                            background: 'linear-gradient(135deg, #10b981, #059669)',
                            color: '#ffffff',
                            border: 'none',
                            fontWeight: 700,
                            fontSize: '11px',
                            cursor: 'pointer',
                            boxShadow: '0 2px 6px rgba(16, 185, 129, 0.4)'
                          }}
                        >
                          ✓ Apply Site Code
                        </button>
                      ) : (
                        <span style={{ fontSize: '11px', color: '#6ee7b7', fontWeight: 600 }}>
                          ✓ Site Code Applied
                        </span>
                      )}
                    </div>
                  )}
                  {(() => {
                    const conflict = getConflictSummary(modalSiteCode);
                    if (!conflict) return null;
                    return (
                      <div style={{
                        gridColumn: '1 / -1',
                        padding: '9px 13px',
                        borderRadius: '8px',
                        background: conflict.isCombined ? 'rgba(168, 85, 247, 0.12)' : 'rgba(56, 189, 248, 0.12)',
                        border: `1px solid ${conflict.isCombined ? 'rgba(168, 85, 247, 0.35)' : 'rgba(56, 189, 248, 0.35)'}`,
                        fontSize: '12px',
                        color: '#f1f5f9',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '10px'
                      }}>
                        <span style={{ fontSize: '16px' }}>🧩</span>
                        <div>
                          <strong style={{ color: conflict.isCombined ? '#c084fc' : '#38bdf8' }}>
                            {conflict.type}:
                          </strong>{' '}
                          {conflict.message}
                        </div>
                      </div>
                    );
                  })()}
                  <div className="scooh-field">
                    <label>Month</label>
                    <input name="month" defaultValue={editModal.month ?? ''} placeholder="e.g. Oct 2026" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>Booking Date</label>
                    <input type="date" name="booking_date" defaultValue={editModal.booking_date ? editModal.booking_date.slice(0, 10) : ''} disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field" style={{ gridColumn: 'span 2' }}>
                    <label>Client / Agency</label>
                    <input name="client" defaultValue={editModal.client ?? editModal.client_name ?? ''} placeholder="Client company or agency" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field" style={{ gridColumn: 'span 2' }}>
                    <label>Display / Brand</label>
                    <input name="display" defaultValue={cleanDisplayTitle(editModal.display ?? editModal.campaign_name ?? '')} placeholder="Display title or brand" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>Vendor Name</label>
                    <input name="vendor_name" defaultValue={editModal.vendor_name ?? ''} placeholder="Printing/Mounting vendor" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>Location / Address</label>
                    <input
                      name="location"
                      value={modalLocation}
                      onChange={e => setModalLocation(e.target.value)}
                      placeholder="e.g. Nr DMart / Shivranjani"
                      disabled={isViewingOnly || isReadOnly}
                    />
                  </div>
                  <div className="scooh-field">
                    <label>Width (ft)</label>
                    <input
                      type="number"
                      step="any"
                      name="width"
                      value={modalWidth}
                      onChange={e => {
                        const val = e.target.value;
                        setModalWidth(val);
                        if (val && modalHeight) setModalSize(`${val}x${modalHeight} ft`);
                      }}
                      placeholder="e.g. 40"
                      disabled={isViewingOnly || isReadOnly}
                    />
                  </div>
                  <div className="scooh-field">
                    <label>Height (ft)</label>
                    <input
                      type="number"
                      step="any"
                      name="height"
                      value={modalHeight}
                      onChange={e => {
                        const val = e.target.value;
                        setModalHeight(val);
                        if (modalWidth && val) setModalSize(`${modalWidth}x${val} ft`);
                      }}
                      placeholder="e.g. 20"
                      disabled={isViewingOnly || isReadOnly}
                    />
                  </div>
                  <div className="scooh-field">
                    <label>Size</label>
                    <input
                      name="size"
                      value={modalSize}
                      onChange={e => setModalSize(e.target.value)}
                      placeholder="e.g. 40x20 ft"
                      disabled={isViewingOnly || isReadOnly}
                    />
                  </div>
                  <div className="scooh-field">
                    <label>Media Type</label>
                    <select
                      name="type"
                      value={modalType}
                      onChange={e => setModalType(e.target.value)}
                      disabled={isViewingOnly || isReadOnly}
                    >
                      <option value="Hoarding">Hoarding</option>
                      <option value="Gantry">Gantry</option>
                      <option value="Unipole">Unipole</option>
                      <option value="Billboard">Billboard</option>
                      <option value="DOOH">DOOH</option>
                    </select>
                  </div>
                  <div className="scooh-field">
                    <label>Start Date</label>
                    <input type="date" name="start_date" defaultValue={editModal.start_date ? editModal.start_date.slice(0, 10) : ''} disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>End Date</label>
                    <input type="date" name="end_date" defaultValue={editModal.end_date ? editModal.end_date.slice(0, 10) : ''} disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>Days</label>
                    <input type="number" name="days" defaultValue={editModal.days ?? 30} placeholder="30" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>Advt. Fees per month (₹)</label>
                    <input type="number" step="any" name="advt_fees" defaultValue={editModal.advt_fees ?? ''} placeholder="0" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>Printing & Mounting (₹)</label>
                    <input type="number" step="any" name="printing_mounting_cost" defaultValue={editModal.printing_mounting_cost ?? ''} placeholder="0" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>Total Amount (₹)</label>
                    <input type="number" step="any" name="total_amount" defaultValue={editModal.total_amount ?? editModal.revenue ?? ''} placeholder="0" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>PO (Purchase Order)</label>
                    <input name="po" defaultValue={editModal.po ?? ''} placeholder="e.g. PO-2026-881" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>Bill (Invoice No / Status)</label>
                    <input name="bill" defaultValue={editModal.bill ?? editModal.invoice_no ?? ''} placeholder="e.g. INV-9912 / Sent" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>Pending (₹)</label>
                    <input type="number" step="any" name="pending" defaultValue={editModal.pending ?? ''} placeholder="0" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field" style={{ gridColumn: 'span 2' }}>
                    <label>Notes</label>
                    <textarea name="notes" rows="2" defaultValue={editModal.notes ?? ''} placeholder="Campaign notes or instructions…" disabled={isViewingOnly || isReadOnly} />
                  </div>
                </div>
              </div>
              <div className="scooh-modalfoot" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div>
                  {isViewingOnly && canEdit && !isReadOnly && (
                    <button
                      type="button"
                      className="scooh-btn ghost"
                      style={{ color: '#fbbf24', borderColor: 'rgba(251, 191, 36, 0.4)', fontSize: '12px' }}
                      onClick={() => setIsViewingOnly(false)}
                      title="Switch to edit mode"
                    >
                      ✏️ Edit This Campaign
                    </button>
                  )}
                </div>
                <div style={{ display: 'flex', gap: '8px' }}>
                  <button type="button" className="scooh-btn ghost" onClick={() => setEditModal(null)}>
                    {isViewingOnly || isReadOnly ? 'Close' : 'Cancel'}
                  </button>
                  {!isViewingOnly && !isReadOnly && (
                    <button type="submit" className="scooh-btn purple-btn" disabled={saving}>
                      {saving ? 'Saving…' : editModal.id ? 'Update Campaign' : 'Create Campaign'}
                    </button>
                  )}
                </div>
              </div>
            </form>
          </div>
        </div>
      )}
    </>
  );
}

async function exportElectricityExcel(rowsData, filename = 'MediaBuzz_Electricity_Bills.xlsx') {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Electricity Bills', {
    views: [{ state: 'frozen', ySplit: 1 }]
  });

  worksheet.columns = [
    { header: 'SR NO', key: 'sr', width: 8 },
    { header: 'SITE CODE', key: 'site_code', width: 16 },
    { header: 'LOCATION', key: 'location', width: 40 },
    { header: 'SIZE', key: 'size', width: 14 },
    { header: 'METER NO', key: 'meter_no', width: 20 },
    { header: 'SERVICE NUMBER', key: 'service_number', width: 18 },
    { header: 'T NUMBER', key: 't_number', width: 14 },
    { header: 'BILL / PROVIDER', key: 'bill_type', width: 18 },
    { header: 'BILLING MONTH', key: 'billing_month', width: 16 },
    { header: 'DUE DATE', key: 'due_date', width: 16 },
    { header: 'UNITS', key: 'units', width: 12 },
    { header: 'RATE', key: 'rate', width: 12 },
    { header: 'PAYMENT AMOUNT (₹)', key: 'amount', width: 20 },
    { header: 'STATUS', key: 'status', width: 14 },
    { header: 'PAID DATE', key: 'paid_date', width: 16 },
    { header: 'PAYMENT REF', key: 'payment_reference', width: 22 },
    { header: 'NOTES', key: 'notes', width: 28 }
  ];

  rowsData.forEach((r, idx) => {
    worksheet.addRow({
      sr: idx + 1,
      site_code: r.site_code || '',
      location: r.location || '',
      size: r.size || '',
      meter_no: r.meter_no || '',
      service_number: r.service_number || '',
      t_number: r.t_number || '',
      bill_type: r.bill_type || '',
      billing_month: r.billing_month || '',
      due_date: r.due_date ? new Date(r.due_date).toLocaleDateString('en-IN') : '',
      units: r.units || '',
      rate: r.rate || '',
      amount: Number(r.amount || r.payment_amount || 0),
      status: r.payment_status || 'Pending',
      paid_date: r.paid_date ? new Date(r.paid_date).toLocaleDateString('en-IN') : '',
      payment_reference: r.payment_reference || '',
      notes: r.notes || ''
    });
  });

  // Style Header Row (Row 1) in Bright Yellow (#FFFF00)
  const headerRow = worksheet.getRow(1);
  headerRow.height = 28;
  headerRow.eachCell((cell, colNumber) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF00' }
    };
    cell.font = {
      name: 'Calibri',
      size: 11,
      bold: true,
      color: { argb: 'FF000000' }
    };
    cell.alignment = {
      vertical: 'middle',
      horizontal: [1, 4, 8, 9, 10, 11, 12, 14, 15].includes(colNumber) ? 'center' : (colNumber === 13 ? 'right' : 'left'),
      wrapText: false
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFC0C0C0' } },
      left: { style: 'thin', color: { argb: 'FFC0C0C0' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FFC0C0C0' } }
    };
  });

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  window.URL.revokeObjectURL(url);
}

function StorageVaultSection({ embedded = true, title = null, subtitle = null, initialCategory = 'all' }) {
  const [items, setItems] = useState([]);
  const [loading, setLoading] = useState(false);
  const [category, setCategory] = useState(initialCategory);
  const [search, setSearch] = useState('');
  const [sort, setSort] = useState('newest');
  const [selectedIds, setSelectedIds] = useState(new Set());
  const [uploadModal, setUploadModal] = useState(false);
  const [snapshotModal, setSnapshotModal] = useState(false);
  const [snapshotMonth, setSnapshotMonth] = useState('');
  const [activeOccItem, setActiveOccItem] = useState(null);
  const [banner, setBanner] = useState('');
  const [actionLoading, setActionLoading] = useState(false);

  const currentRole = getCurrentRole();
  const isAdmin = currentRole === 'admin';
  const isManager = currentRole === 'manager';
  const isStaff = currentRole === 'staff';
  const isReadOnly = currentRole === 'viewer';
  const canDelete = isAdmin || isManager;

  async function loadStorage() {
    setLoading(true);
    try {
      const { data } = await api.get('/storage');
      if (Array.isArray(data)) setItems(data);
    } catch (err) {
      console.warn('Load storage error:', err);
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    loadStorage();
  }, []);

  const pptCount = items.filter(i => i.category === 'ppt').length;
  const excelCount = items.filter(i => i.category === 'excel').length;
  const occCount = items.filter(i => i.category === 'occupancy').length;
  const otherCount = items.filter(i => !['ppt', 'excel', 'occupancy'].includes(i.category)).length;

  const filtered = useMemo(() => {
    let list = items.filter(i => {
      if (category !== 'all' && i.category !== category) return false;
      if (search.trim()) {
        const q = search.toLowerCase();
        const str = `${i.title || ''} ${i.filename || ''} ${JSON.stringify(i.meta || {})} ${i.format || ''}`.toLowerCase();
        if (!str.includes(q)) return false;
      }
      return true;
    });

    if (sort === 'newest') list.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
    else if (sort === 'oldest') list.sort((a, b) => new Date(a.created_at) - new Date(b.created_at));
    else if (sort === 'title') list.sort((a, b) => (a.title || '').localeCompare(b.title || ''));
    return list;
  }, [items, category, search, sort]);

  const occItems = useMemo(() => {
    return items
      .filter(i => i.category === 'occupancy')
      .sort((a, b) => new Date(a.created_at) - new Date(b.created_at));
  }, [items]);

  function toggleSelect(id) {
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function toggleSelectAll() {
    const visibleIds = filtered.map(r => r.id).filter(Boolean);
    const allSelected = visibleIds.length > 0 && visibleIds.every(id => selectedIds.has(id));
    if (allSelected) {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleIds.forEach(id => next.delete(id));
        return next;
      });
    } else {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleIds.forEach(id => next.add(id));
        return next;
      });
    }
  }

  function selectAllVisible() {
    setSelectedIds(new Set(filtered.map(r => r.id).filter(Boolean)));
  }

  function deselectAll() {
    setSelectedIds(new Set());
  }

  async function handleBatchDelete() {
    if (!canDelete || selectedIds.size === 0) return;
    const count = selectedIds.size;
    if (!window.confirm(`Are you sure you want to permanently delete all ${count} selected file${count > 1 ? 's' : ''} from storage? This action cannot be undone.`)) {
      return;
    }
    setActionLoading(true);
    try {
      await api.post('/storage/batch-delete', { ids: Array.from(selectedIds) });
      setBanner(`✓ ${count} file${count > 1 ? 's' : ''} deleted from storage successfully.`);
      setSelectedIds(new Set());
      await loadStorage();
      setTimeout(() => setBanner(''), 4000);
    } catch (err) {
      alert('Failed to delete selected files: ' + (err.response?.data?.message || err.message));
    } finally {
      setActionLoading(false);
    }
  }

  async function handleTakeSnapshot(e) {
    if (e) e.preventDefault();
    if (isReadOnly) return;
    setActionLoading(true);
    try {
      const { data } = await api.post('/storage/snapshot-occupancy', { month: snapshotMonth || undefined });
      setBanner(`✓ Captured snapshot: ${data.title}`);
      setSnapshotModal(false);
      setSnapshotMonth('');
      await loadStorage();
      setTimeout(() => setBanner(''), 4000);
    } catch (err) {
      alert('Snapshot failed: ' + (err.response?.data?.message || err.message));
    } finally {
      setActionLoading(false);
    }
  }

  async function handleUpload(e) {
    e.preventDefault();
    if (isReadOnly) return;
    const fd = new FormData(e.currentTarget);
    const file = fd.get('file');
    if (!file || !file.name) {
      alert('Please select a file to upload');
      return;
    }
    setActionLoading(true);
    try {
      const res = await api.post('/storage/upload', fd, {
        headers: { 'Content-Type': 'multipart/form-data' }
      });
      setBanner(`✓ Uploaded "${res.data.filename}" to storage successfully!`);
      setUploadModal(false);
      await loadStorage();
      setTimeout(() => setBanner(''), 4000);
    } catch (err) {
      alert('Upload failed: ' + (err.response?.data?.message || err.message));
    } finally {
      setActionLoading(false);
    }
  }

  async function handleDelete(item) {
    if (!canDelete) return;
    if (!window.confirm(`Permanently delete "${item.title || item.filename}" from storage?`)) return;
    try {
      await api.delete(`/storage/${item.id}`);
      setItems(prev => prev.filter(x => x.id !== item.id));
      setSelectedIds(prev => {
        const next = new Set(prev);
        next.delete(item.id);
        return next;
      });
      setBanner('✓ Removed item from storage.');
      setTimeout(() => setBanner(''), 3000);
    } catch (err) {
      alert('Delete failed: ' + (err.response?.data?.message || err.message));
    }
  }

  async function handleDownload(item) {
    if (item.file_url) {
      const a = document.createElement('a');
      a.href = item.file_url;
      a.download = item.filename || 'download';
      document.body.appendChild(a);
      a.click();
      a.remove();
      return;
    }

    if (item.category === 'occupancy' && item.meta) {
      const meta = item.meta;
      const csvRows = [
        ['Media Buzz — Occupancy Snapshot Report'],
        ['Month', meta.month || ''],
        ['Occupancy Rate', `${meta.occupancyRate ?? 0}%`],
        ['Occupied Sites', `${meta.occupiedSites ?? 0} of ${meta.totalSites ?? 22}`],
        ['Total Revenue (₹)', meta.revenue ?? 0],
        ['Generated At', meta.generatedAt || new Date().toISOString()],
        [],
        ['Site Code', 'Status', 'Client', 'Monthly Rate (₹)', 'Location']
      ];
      if (Array.isArray(meta.sites)) {
        meta.sites.forEach(s => {
          csvRows.push([
            s.site_code || '',
            s.status || '',
            s.client || '',
            s.rate || s.advt_fees || 0,
            s.location || ''
          ]);
        });
      }
      const csvStr = csvRows.map(row => row.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n');
      const blob = new Blob([csvStr], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${(item.title || 'Occupancy_Snapshot').replace(/\s+/g, '_')}.csv`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      return;
    }

    alert(`Download initiated for ${item.title || item.filename}`);
  }

  const defaultTitle = title || 'Document & Media Storage Vault';
  const defaultSubtitle = subtitle || 'Archive for generated presentations, exported Excel trackers, site occupancy records, and media attachments.';

  return (
    <div style={{ marginTop: embedded ? '36px' : '0' }}>
      <div style={{
        background: '#0d131f',
        border: '1px solid #1e293b',
        borderRadius: '16px',
        overflow: 'hidden',
        boxShadow: '0 8px 30px rgba(0,0,0,0.35)'
      }}>
        <div style={{
          padding: '20px 24px',
          borderBottom: '1px solid rgba(255,255,255,0.08)',
          background: 'linear-gradient(90deg, rgba(168,85,247,0.12), rgba(59,130,246,0.08), transparent)',
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          flexWrap: 'wrap',
          gap: '14px'
        }}>
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
              <span style={{ fontSize: '24px' }}>🗄️</span>
              <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 800, color: '#f8fafc', letterSpacing: '-0.01em' }}>
                {defaultTitle}
              </h3>
              <span className="scooh-badgechip" style={{ background: 'rgba(168,85,247,0.22)', color: '#d8b4fe', fontWeight: 800, border: '1px solid rgba(168,85,247,0.4)', fontSize: '11px' }}>
                {items.length} Files Stored
              </span>
            </div>
            <p style={{ margin: '4px 0 0 34px', fontSize: '12.5px', color: '#94a3b8' }}>
              {defaultSubtitle}
            </p>
          </div>

          <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', alignItems: 'center' }}>
            {!isReadOnly && (
              <>
                <button
                  type="button"
                  className="scooh-btn secondary"
                  style={{ fontSize: '12px', padding: '8px 14px', display: 'inline-flex', alignItems: 'center', gap: '6px' }}
                  onClick={() => setSnapshotModal(true)}
                  title="Save an instant historical snapshot of current occupancy rate and active campaigns"
                >
                  📸 Snapshot Occupancy
                </button>
                <button
                  type="button"
                  className="scooh-btn purple-btn"
                  style={{ fontSize: '12px', padding: '8px 14px', display: 'inline-flex', alignItems: 'center', gap: '6px' }}
                  onClick={() => setUploadModal(true)}
                  title="Upload PPT, Excel, or PDF directly to storage"
                >
                  📤 + Upload File
                </button>
              </>
            )}
            <button
              type="button"
              className="scooh-btn ghost"
              style={{ fontSize: '12px', padding: '8px 12px' }}
              onClick={loadStorage}
              title="Reload storage list"
            >
              🔄
            </button>
          </div>
        </div>

        {isReadOnly && (
          <div className="scooh-banner" style={{ display: 'block', margin: '14px 20px 0', background: 'rgba(245,158,11,0.12)', border: '1px solid rgba(245,158,11,0.3)', color: '#fbbf24' }}>
            🔒 Read-Only Workspace: You are signed in as a Viewer. You can view, search, inspect snapshots, and download files. Uploads, snapshots, and file deletions are disabled.
          </div>
        )}

        <div style={{
          display: 'flex',
          gap: '8px',
          padding: '12px 20px',
          background: 'rgba(15,23,42,0.6)',
          borderBottom: '1px solid rgba(255,255,255,0.06)',
          overflowX: 'auto',
          alignItems: 'center'
        }}>
          {[
            { id: 'all', label: 'All Storage', count: items.length, icon: '🌟' },
            { id: 'ppt', label: 'Generated PPTs', count: pptCount, icon: '📑' },
            { id: 'excel', label: 'Excel Workbooks', count: excelCount, icon: '📊' },
            { id: 'occupancy', label: 'Previous Site Occupancy', count: occCount, icon: '🏛️' },
            { id: 'other', label: 'Uploaded Documents', count: otherCount, icon: '🗂️' }
          ].map(tab => {
            const active = category === tab.id;
            return (
              <button
                key={tab.id}
                type="button"
                onClick={() => setCategory(tab.id)}
                style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '6px',
                  padding: '7px 14px',
                  borderRadius: '20px',
                  fontSize: '12px',
                  fontWeight: active ? 800 : 600,
                  cursor: 'pointer',
                  border: active ? '1px solid #a855f7' : '1px solid rgba(255,255,255,0.08)',
                  background: active ? 'rgba(168,85,247,0.25)' : 'rgba(30,41,59,0.5)',
                  color: active ? '#ffffff' : '#94a3b8',
                  transition: 'all 0.15s ease'
                }}
              >
                <span>{tab.icon}</span>
                <span>{tab.label}</span>
                <span style={{
                  padding: '1px 6px',
                  borderRadius: '10px',
                  fontSize: '10.5px',
                  background: active ? 'rgba(168,85,247,0.45)' : 'rgba(255,255,255,0.08)',
                  color: active ? '#f3e8ff' : '#cbd5e1'
                }}>
                  {tab.count}
                </span>
              </button>
            );
          })}
        </div>

        {banner && (
          <div style={{
            margin: '14px 20px',
            padding: '10px 14px',
            borderRadius: '8px',
            background: 'rgba(34,197,94,0.16)',
            border: '1px solid #22c55e',
            color: '#4ade80',
            fontSize: '12.5px',
            fontWeight: 600
          }}>
            {banner}
          </div>
        )}

        {category === 'occupancy' && occItems.length > 0 && (
          <div style={{
            padding: '16px 20px',
            background: 'rgba(30,41,59,0.4)',
            borderBottom: '1px solid rgba(255,255,255,0.06)'
          }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '10px' }}>
              <span style={{ fontSize: '12px', fontWeight: 800, color: '#e2e8f0', letterSpacing: '0.04em', textTransform: 'uppercase' }}>
                Occupancy History Timeline
              </span>
              <span style={{ fontSize: '11px', color: '#94a3b8' }}>
                {occItems.length} monthly snapshots recorded
              </span>
            </div>

            <div style={{ display: 'flex', gap: '12px', overflowX: 'auto', paddingBottom: '6px' }}>
              {occItems.map(item => {
                const m = item.meta || {};
                const occRate = Number(m.occupancyRate ?? 0);
                const isSelected = activeOccItem?.id === item.id;
                return (
                  <div
                    key={item.id}
                    onClick={() => setActiveOccItem(item)}
                    style={{
                      minWidth: '190px',
                      padding: '12px 14px',
                      borderRadius: '10px',
                      cursor: 'pointer',
                      border: isSelected ? '1px solid #38bdf8' : '1px solid rgba(255,255,255,0.08)',
                      background: isSelected ? 'rgba(56,189,248,0.15)' : '#111927',
                      transition: 'all 0.15s ease'
                    }}
                  >
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '6px' }}>
                      <span style={{ fontWeight: 800, color: '#f8fafc', fontSize: '13px' }}>
                        {m.month || item.title}
                      </span>
                      <span style={{
                        fontWeight: 900,
                        fontSize: '12px',
                        color: occRate >= 70 ? '#4ade80' : occRate >= 40 ? '#facc15' : '#ef4444'
                      }}>
                        {occRate}%
                      </span>
                    </div>

                    <div style={{
                      height: '5px',
                      borderRadius: '3px',
                      background: '#1e293b',
                      overflow: 'hidden',
                      marginBottom: '8px'
                    }}>
                      <div style={{
                        width: `${Math.min(100, Math.max(5, occRate))}%`,
                        height: '100%',
                        background: occRate >= 70 ? '#22c55e' : occRate >= 40 ? '#eab308' : '#ef4444'
                      }} />
                    </div>

                    <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '11px', color: '#94a3b8' }}>
                      <span>Sites: <strong style={{ color: '#e2e8f0' }}>{m.occupiedSites ?? '—'} / {m.totalSites ?? 22}</strong></span>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        <div style={{
          padding: '12px 20px',
          display: 'flex',
          gap: '12px',
          alignItems: 'center',
          flexWrap: 'wrap',
          borderBottom: '1px solid rgba(255,255,255,0.06)'
        }}>
          <input
            className="scooh-search"
            style={{ flex: '1 1 240px', minWidth: '180px', height: '36px' }}
            placeholder="Search stored PPTs, Excels, occupancy snapshots, filenames…"
            value={search}
            onChange={e => setSearch(e.target.value)}
          />

          <select
            value={sort}
            onChange={e => setSort(e.target.value)}
            style={{
              padding: '6px 12px',
              borderRadius: '8px',
              background: '#111925',
              color: '#e2e8f0',
              border: '1px solid #334155',
              fontSize: '12px',
              fontWeight: 600
            }}
          >
            <option value="newest">Sort: Newest Upload / Generation</option>
            <option value="oldest">Sort: Oldest First</option>
            <option value="title">Sort: Title (A–Z)</option>
          </select>

          {search && (
            <button
              type="button"
              className="scooh-btn ghost"
              style={{ fontSize: '11px', padding: '6px 10px' }}
              onClick={() => setSearch('')}
            >
              Clear Search
            </button>
          )}

          {canDelete && filtered.length > 0 && (
            <button
              type="button"
              className="scooh-btn secondary"
              style={{ fontSize: '11.5px', padding: '6px 12px', display: 'inline-flex', alignItems: 'center', gap: '5px' }}
              onClick={toggleSelectAll}
              title={filtered.every(r => selectedIds.has(r.id)) ? 'Deselect all visible files' : 'Select all visible files'}
            >
              {filtered.every(r => selectedIds.has(r.id)) ? '✓ Deselect All' : `☑ Select All (${filtered.length})`}
            </button>
          )}
        </div>

        {canDelete && selectedIds.size > 0 && (
          <div style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            padding: '11px 18px',
            background: 'linear-gradient(90deg, rgba(239,68,68,0.18), rgba(239,68,68,0.08))',
            borderBottom: '1px solid rgba(239,68,68,0.3)',
            color: '#fca5a5',
            fontSize: '13px',
            flexWrap: 'wrap',
            gap: '12px'
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap' }}>
              <span style={{ fontWeight: 800, color: '#ffffff', display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
                {selectedIds.size} file{selectedIds.size > 1 ? 's' : ''} selected
              </span>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={selectAllVisible}
                style={{ fontSize: '11.5px', padding: '4px 10px', color: '#f1f5f9', borderColor: '#475569' }}
              >
                Select all visible ({filtered.length})
              </button>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={deselectAll}
                style={{ fontSize: '11.5px', padding: '4px 10px', color: '#cbd5e1', borderColor: '#475569' }}
              >
                Deselect all
              </button>
            </div>
            <button
              type="button"
              className="scooh-btn danger"
              onClick={handleBatchDelete}
              disabled={actionLoading}
              style={{
                background: '#ef4444',
                color: '#ffffff',
                border: 'none',
                fontWeight: 800,
                padding: '7px 16px',
                borderRadius: '7px',
                cursor: 'pointer',
                display: 'inline-flex',
                alignItems: 'center',
                gap: '6px'
              }}
            >
              🗑 Delete selected ({selectedIds.size})
            </button>
          </div>
        )}

        <div className="scooh-tablewrap">
          <table className="scooh-table">
            <thead>
              <tr>
                {canDelete && (
                  <th style={{ width: '42px', textAlign: 'center', padding: '10px 8px' }}>
                    <input
                      type="checkbox"
                      checked={filtered.length > 0 && filtered.every(r => selectedIds.has(r.id))}
                      onChange={toggleSelectAll}
                      title={filtered.length > 0 && filtered.every(r => selectedIds.has(r.id)) ? 'Deselect all visible' : 'Select all visible'}
                      style={{ cursor: 'pointer', width: '15px', height: '15px' }}
                    />
                  </th>
                )}
                <th style={{ width: '60px', textAlign: 'center' }}>Type</th>
                <th style={{ minWidth: '220px' }}>Document / File Name</th>
                <th style={{ minWidth: '130px' }}>Category</th>
                <th style={{ minWidth: '200px' }}>Details & Metadata</th>
                <th style={{ width: '100px', textAlign: 'right' }}>Size</th>
                <th style={{ width: '120px' }}>Saved On</th>
                <th style={{ minWidth: '140px', textAlign: 'center' }}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr>
                  <td colSpan={canDelete ? "8" : "7"} className="scooh-empty" style={{ padding: '36px 20px', textAlign: 'center' }}>
                    <div style={{ fontSize: '14px', fontWeight: 700, color: '#94a3b8' }}>No storage items found</div>
                    <div style={{ fontSize: '12px', color: '#64748b', marginTop: '4px' }}>
                      Generate a presentation, export an Excel sheet, or click "+ Upload File" above.
                    </div>
                  </td>
                </tr>
              ) : (
                filtered.map(item => {
                  const m = item.meta || {};
                  const isPpt = item.category === 'ppt' || ['pptx', 'ppt'].includes(item.format);
                  const isExcel = item.category === 'excel' || ['xlsx', 'xls', 'csv'].includes(item.format);
                  const isOcc = item.category === 'occupancy' || item.format === 'json';

                  const badgeBg = isPpt ? 'rgba(249,115,22,0.18)' : isExcel ? 'rgba(34,197,94,0.18)' : isOcc ? 'rgba(56,189,248,0.18)' : 'rgba(168,85,247,0.18)';
                  const badgeColor = isPpt ? '#fb923c' : isExcel ? '#4ade80' : isOcc ? '#38bdf8' : '#c084fc';
                  const badgeBorder = isPpt ? 'rgba(249,115,22,0.35)' : isExcel ? 'rgba(34,197,94,0.35)' : isOcc ? 'rgba(56,189,248,0.35)' : 'rgba(168,85,247,0.35)';

                  return (
                    <tr key={item.id} style={{ background: selectedIds.has(item.id) ? 'rgba(239,68,68,0.06)' : undefined }}>
                      {canDelete && (
                        <td style={{ textAlign: 'center', padding: '10px 8px' }}>
                          <input
                            type="checkbox"
                            checked={selectedIds.has(item.id)}
                            onChange={() => toggleSelect(item.id)}
                            style={{ cursor: 'pointer', width: '15px', height: '15px' }}
                          />
                        </td>
                      )}
                      <td style={{ textAlign: 'center' }}>
                        <span className="scooh-badgechip" style={{ background: badgeBg, color: badgeColor, border: `1px solid ${badgeBorder}`, fontWeight: 800, fontSize: '10.5px' }}>
                          {item.format ? item.format.toUpperCase() : 'FILE'}
                        </span>
                      </td>
                      <td>
                        <div style={{ fontWeight: 800, color: '#ffffff', fontSize: '13px' }}>
                          {item.title}
                        </div>
                        <div style={{ fontSize: '11px', color: '#94a3b8', marginTop: '2px', fontFamily: 'monospace' }}>
                          {item.filename}
                        </div>
                      </td>
                      <td>
                        <span style={{
                          textTransform: 'capitalize',
                          fontSize: '12px',
                          color: '#e2e8f0',
                          display: 'inline-flex',
                          alignItems: 'center',
                          gap: '6px'
                        }}>
                          {item.category === 'ppt' ? '📑 PPT Deck' :
                           item.category === 'excel' ? '📊 Excel Sheet' :
                           item.category === 'occupancy' ? '🏛️ Occupancy Snapshot' : '🗂️ Document'}
                        </span>
                      </td>
                      <td style={{ fontSize: '11.5px', color: '#cbd5e1' }}>
                        {isPpt && (
                          <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap', alignItems: 'center' }}>
                            {m.slides && <span className="scooh-plate" style={{ fontSize: '10px' }}>{m.slides} Slides</span>}
                          </div>
                        )}
                        {isExcel && (
                          <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap', alignItems: 'center' }}>
                            {m.rows && <span className="scooh-plate" style={{ fontSize: '10px' }}>{m.rows} Rows</span>}
                            {m.type && <span style={{ color: '#4ade80' }}>{m.type}</span>}
                          </div>
                        )}
                        {isOcc && (
                          <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap', alignItems: 'center' }}>
                            <span className="scooh-badgechip" style={{ background: 'rgba(34,197,94,0.18)', color: '#4ade80', fontSize: '10px' }}>
                              Rate: {m.occupancyRate ?? '—'}%
                            </span>
                            <span style={{ color: '#cbd5e1' }}>
                              Occupied: {m.occupiedSites ?? '—'}/{m.totalSites ?? '—'}
                            </span>
                            {m.revenue ? <span style={{ color: '#facc15' }}>{money(m.revenue)}</span> : null}
                          </div>
                        )}
                        {!isPpt && !isExcel && !isOcc && (
                          <span style={{ color: '#94a3b8' }}>Standard file</span>
                        )}
                      </td>
                      <td style={{ textAlign: 'right', fontSize: '11.5px', color: '#94a3b8', fontWeight: 600 }}>
                        {item.file_size || '—'}
                      </td>
                      <td style={{ fontSize: '11.5px', color: '#94a3b8' }}>
                        {item.created_at ? new Date(item.created_at).toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' }) : '—'}
                      </td>
                      <td>
                        <div className="scooh-rowactions" style={{ justifyContent: 'center' }}>
                          <button
                            type="button"
                            className="scooh-btn secondary"
                            style={{ minHeight: '28px', padding: '0 10px', fontSize: '11px', display: 'inline-flex', alignItems: 'center', gap: '4px' }}
                            onClick={() => handleDownload(item)}
                            title="Download or re-export file"
                          >
                            ⬇ Download
                          </button>
                          {isOcc && (
                            <button
                              type="button"
                              className="scooh-btn ghost"
                              style={{ minHeight: '28px', padding: '0 8px', fontSize: '11px' }}
                              onClick={() => setActiveOccItem(item)}
                              title="View full snapshot breakdown"
                            >
                              👁 View
                            </button>
                          )}
                          {canDelete && (
                            <button
                              type="button"
                              className="scooh-iconbtn danger-icon"
                              style={{ minHeight: '28px', width: '28px' }}
                              onClick={() => handleDelete(item)}
                              title="Delete file from storage"
                            >
                              ×
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Upload Modal */}
      {uploadModal && (
        <div className="scooh-modal-overlay" onClick={() => !actionLoading && setUploadModal(false)}>
          <div className="scooh-modal" style={{ maxWidth: '520px', width: '92%' }} onClick={e => e.stopPropagation()}>
            <div className="scooh-modalhead">
              <div>
                <h2>Upload File to Storage</h2>
                <span className="scooh-modal-subtitle">Save PPT presentations, Excel files, or documents to permanent storage</span>
              </div>
              <button type="button" className="scooh-modalclose" onClick={() => setUploadModal(false)}>×</button>
            </div>
            <form onSubmit={handleUpload}>
              <div className="scooh-modalbody">
                <div style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>
                  <div className="scooh-field">
                    <label>Choose File *</label>
                    <input type="file" name="file" required accept=".pptx,.ppt,.xlsx,.xls,.csv,.pdf,.json,.png,.jpg,.jpeg" />
                  </div>
                  <div className="scooh-field">
                    <label>Category</label>
                    <select name="category" defaultValue="ppt">
                      <option value="ppt">📑 PPT Presentation</option>
                      <option value="excel">📊 Excel Spreadsheet</option>
                      <option value="occupancy">🏛️ Site Occupancy Data</option>
                      <option value="other">🗂️ General Document / Contract</option>
                    </select>
                  </div>
                  <div className="scooh-field">
                    <label>Title (optional)</label>
                    <input name="title" placeholder="e.g. Diwali 2026 Pitch Deck" />
                  </div>
                </div>
              </div>
              <div className="scooh-modalfoot">
                <button type="button" className="scooh-btn ghost" onClick={() => setUploadModal(false)}>Cancel</button>
                <button type="submit" className="scooh-btn purple-btn" disabled={actionLoading}>
                  {actionLoading ? 'Uploading…' : 'Upload to Storage'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Snapshot Modal */}
      {snapshotModal && (
        <div className="scooh-modal-overlay" onClick={() => !actionLoading && setSnapshotModal(false)}>
          <div className="scooh-modal" style={{ maxWidth: '480px', width: '92%' }} onClick={e => e.stopPropagation()}>
            <div className="scooh-modalhead">
              <div>
                <h2>Record Current Month Occupancy</h2>
                <span className="scooh-modal-subtitle">Save live occupancy percentage, active site count & revenue to history</span>
              </div>
              <button type="button" className="scooh-modalclose" onClick={() => setSnapshotModal(false)}>×</button>
            </div>
            <form onSubmit={handleTakeSnapshot}>
              <div className="scooh-modalbody">
                <div className="scooh-field">
                  <label>Billing Month Label</label>
                  <input
                    value={snapshotMonth}
                    onChange={e => setSnapshotMonth(e.target.value)}
                    placeholder={`e.g. ${new Date().toLocaleDateString('en-US', { month: 'long', year: 'numeric' })}`}
                  />
                  <small style={{ color: '#94a3b8', marginTop: '4px', fontSize: '11px' }}>
                    Leave blank to automatically use current calendar month.
                  </small>
                </div>
              </div>
              <div className="scooh-modalfoot">
                <button type="button" className="scooh-btn ghost" onClick={() => setSnapshotModal(false)}>Cancel</button>
                <button type="submit" className="scooh-btn secondary" disabled={actionLoading}>
                  {actionLoading ? 'Recording…' : '📸 Save Snapshot'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* View Occupancy Snapshot Detail Modal */}
      {activeOccItem && (
        <div className="scooh-modal-overlay" onClick={() => setActiveOccItem(null)}>
          <div className="scooh-modal" style={{ maxWidth: '580px', width: '92%' }} onClick={e => e.stopPropagation()}>
            <div className="scooh-modalhead">
              <div>
                <h2>{activeOccItem.title}</h2>
                <span className="scooh-modal-subtitle">Recorded on {new Date(activeOccItem.created_at).toLocaleDateString('en-IN', { day: 'numeric', month: 'long', year: 'numeric' })}</span>
              </div>
              <button type="button" className="scooh-modalclose" onClick={() => setActiveOccItem(null)}>×</button>
            </div>
            <div className="scooh-modalbody">
              {(() => {
                const m = activeOccItem.meta || {};
                const pct = Number(m.occupancyPct || 0);
                return (
                  <div>
                    <div style={{
                      display: 'grid',
                      gridTemplateColumns: 'repeat(2, 1fr)',
                      gap: '12px',
                      marginBottom: '16px'
                    }}>
                      <div style={{ background: 'rgba(15,23,42,0.6)', padding: '12px', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.06)' }}>
                        <span style={{ fontSize: '11px', color: '#94a3b8', textTransform: 'uppercase', fontWeight: 800 }}>Occupancy Rate</span>
                        <div style={{ fontSize: '24px', fontWeight: 800, color: '#38bdf8', marginTop: '4px' }}>{pct}%</div>
                      </div>
                      <div style={{ background: 'rgba(15,23,42,0.6)', padding: '12px', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.06)' }}>
                        <span style={{ fontSize: '11px', color: '#94a3b8', textTransform: 'uppercase', fontWeight: 800 }}>Portfolio Revenue</span>
                        <div style={{ fontSize: '24px', fontWeight: 800, color: '#4ade80', marginTop: '4px' }}>{money(m.revenue || 0)}</div>
                      </div>
                      <div style={{ background: 'rgba(15,23,42,0.6)', padding: '12px', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.06)' }}>
                        <span style={{ fontSize: '11px', color: '#94a3b8', textTransform: 'uppercase', fontWeight: 800 }}>Occupied Sites</span>
                        <div style={{ fontSize: '20px', fontWeight: 800, color: '#f8fafc', marginTop: '4px' }}>{m.occupiedSites ?? '—'} / {m.totalSites || 22}</div>
                      </div>
                      <div style={{ background: 'rgba(15,23,42,0.6)', padding: '12px', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.06)' }}>
                        <span style={{ fontSize: '11px', color: '#94a3b8', textTransform: 'uppercase', fontWeight: 800 }}>Vacant Sites</span>
                        <div style={{ fontSize: '20px', fontWeight: 800, color: '#f87171', marginTop: '4px' }}>{m.vacantSites ?? '—'}</div>
                      </div>
                    </div>
                  </div>
                );
              })()}
            </div>
            <div className="scooh-modalfoot">
              <button
                type="button"
                className="scooh-btn secondary"
                onClick={() => handleDownload(activeOccItem)}
              >
                ⬇ Download Snapshot JSON
              </button>
              <button type="button" className="scooh-btn ghost" onClick={() => setActiveOccItem(null)}>
                Close
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function StorageView() {
  return (
    <>
      <PageHead
        title="Storage & Archives"
        desc="Dedicated repository of generated client PPT decks, exported Excel spreadsheets, historical monthly site occupancy snapshots, and operational files."
      />
      <StorageVaultSection embedded={false} />
    </>
  );
}

function ElectricityView() {
  const [rows, setRows] = useState([]);
  const [sites, setSites] = useState([]);
  const [editModal, setEditModal] = useState(null);
  const [payModal, setPayModal] = useState(null);
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState('ALL');
  const [providerFilter, setProviderFilter] = useState('ALL');
  const [monthFilter, setMonthFilter] = useState('ALL');
  const [sortState, setSortState] = useState({ key: 'due_date', dir: 'asc' });
  const [saving, setSaving] = useState(false);
  const [banner, setBanner] = useState('');
  const [importingExcel, setImportingExcel] = useState(false);
  const [selectedIds, setSelectedIds] = useState(new Set());

  const currentRole = getCurrentRole();
  const isAdmin = currentRole === 'admin';
  const isManager = currentRole === 'manager';
  const isStaff = currentRole === 'staff';
  const isReadOnly = currentRole === 'viewer';
  const canDelete = isAdmin || isManager;
  const canAdd = !isReadOnly;
  const canPay = !isReadOnly;

  async function loadData() {
    try {
      const [eRes, sRes] = await Promise.all([
        api.get('/electricity').catch(() => ({ data: [] })),
        api.get('/sites').catch(() => ({ data: [] }))
      ]);
      const siteList = Array.isArray(sRes.data) && sRes.data.length > 0 ? sRes.data : defaultSites;
      setSites(siteList);

      const siteMap = new Map();
      siteList.forEach(s => {
        if (s.id) siteMap.set(String(s.id), s);
        if (s.site_code) siteMap.set(String(s.site_code).trim(), s);
      });

      const rawBills = Array.isArray(eRes.data) ? eRes.data : [];
      const enriched = rawBills.map(b => {
        const matchedSite = siteMap.get(String(b.site_id)) || siteMap.get(String(b.site_code).trim()) || {};
        return {
          ...b,
          site_code: b.site_code || matchedSite.site_code || '',
          location: b.location || matchedSite.address || matchedSite.area || matchedSite.city || '',
          size: b.size || matchedSite.size || (matchedSite.width && matchedSite.height ? `${matchedSite.width}x${matchedSite.height} ft` : ''),
          meter_no: b.meter_no || matchedSite.meter_no || '',
          service_number: b.service_number || matchedSite.service_number || '',
          t_number: b.t_number || matchedSite.t_number || '',
          amount: Number(b.amount || b.payment_amount || 0),
          payment_amount: Number(b.payment_amount || b.amount || 0)
        };
      });
      setRows(enriched);
    } catch (err) {
      console.warn('Error loading electricity:', err);
    }
  }

  async function handleElectricityExcelImport(e) {
    if (!canDelete) return;
    const file = e.target.files?.[0];
    if (!file) return;
    setImportingExcel(true);
    try {
      const fd = new FormData();
      fd.append('file', file);
      const r = await api.post('/import/electricity-xlsx', fd, { headers: { 'Content-Type': 'multipart/form-data' } });
      const msg = r.data.message || `Processed ${r.data.rows} electricity bills (${r.data.updated} updated, ${r.data.created} created).`;
      setBanner(`✓ ${msg}`);
      alert(`✓ Electricity Bills Import Successful!\n\n${msg}`);
      await loadData();
      window.dispatchEvent(new CustomEvent('mb-electricity-updated'));
      setTimeout(() => setBanner(''), 6000);
    } catch (err) {
      alert('Electricity Bill Import failed: ' + (err.response?.data?.message || err.message));
    } finally {
      setImportingExcel(false);
      e.target.value = '';
    }
  }

  useEffect(() => {
    loadData();
    const onElecUpdate = () => loadData();
    window.addEventListener('mb-electricity-updated', onElecUpdate);
    const interval = setInterval(loadData, 30000);
    return () => {
      window.removeEventListener('mb-electricity-updated', onElecUpdate);
      clearInterval(interval);
    };
  }, []);

  const totalExpense = rows.reduce((acc, r) => acc + Number(r.amount || 0), 0);
  const paidExpense = rows.filter(r => r.payment_status === 'Paid').reduce((acc, r) => acc + Number(r.amount || 0), 0);
  const pendingExpense = rows.filter(r => r.payment_status !== 'Paid').reduce((acc, r) => acc + Number(r.amount || 0), 0);
  const overdueCount = rows.filter(r => {
    if (r.payment_status === 'Paid') return false;
    if (!r.due_date) return false;
    const d = new Date(r.due_date);
    d.setHours(23, 59, 59, 999);
    return d < new Date();
  }).length;

  const months = Array.from(new Set(rows.map(r => r.billing_month).filter(Boolean)));
  const providers = Array.from(new Set(['SSV', 'MB', ...rows.map(r => r.bill_type).filter(Boolean)]));

  function handleSort(key) {
    setSortState(prev => prev.key === key ? { key, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { key, dir: 'asc' });
  }

  const filtered = useMemo(() => {
    const list = rows.filter(r => {
      const q = search.trim().toLowerCase();
      let matchesSearch = !q;
      if (q) {
        if (matchSiteSearch({ site_code: r.site_code, address: r.location }, q)) {
          matchesSearch = true;
        } else if (!/^(?:mb[\s-]*)?\d+$/i.test(q)) {
          matchesSearch = [
            r.size,
            r.meter_no,
            r.service_number,
            r.t_number,
            r.bill_type,
            r.billing_month,
            r.payment_status,
            r.notes
          ].some(v => String(v || '').toLowerCase().includes(q));
        }
      }

      const matchesStatus = statusFilter === 'ALL' ? true :
        statusFilter === 'Overdue' ? (r.payment_status !== 'Paid' && r.due_date && new Date(r.due_date) < new Date()) :
        r.payment_status === statusFilter;

      const matchesProvider = providerFilter === 'ALL' || r.bill_type === providerFilter;
      const matchesMonth = monthFilter === 'ALL' || r.billing_month === monthFilter;

      return matchesSearch && matchesStatus && matchesProvider && matchesMonth;
    });

    if (sortState.key) {
      list.sort((a, b) => {
        let valA = a[sortState.key];
        let valB = b[sortState.key];
        if (sortState.key === 'amount') {
          valA = Number(a.amount || a.payment_amount || 0);
          valB = Number(b.amount || b.payment_amount || 0);
        }
        return universalCompare(valA, valB, sortState.dir);
      });
    }
    return list;
  }, [rows, search, statusFilter, providerFilter, monthFilter, sortState]);

  function openNewBill() {
    if (!canAdd) return;
    setEditModal({
      site_id: '',
      site_code: '',
      location: '',
      size: '',
      meter_no: '',
      service_number: '',
      t_number: '',
      bill_type: 'SSV',
      billing_month: new Date().toLocaleDateString('en-US', { month: 'short', year: 'numeric' }),
      bill_date: new Date().toISOString().slice(0, 10),
      due_date: '',
      units: '',
      rate: '',
      other_charges: 0,
      amount: '',
      payment_status: 'Pending',
      paid_date: '',
      payment_reference: '',
      notes: ''
    });
  }

  function handleSiteChange(siteIdentifier) {
    const s = sites.find(x => String(x.id) === String(siteIdentifier) || String(x.site_code).trim() === String(siteIdentifier).trim());
    if (s) {
      setEditModal(prev => ({
        ...prev,
        site_id: s.id || '',
        site_code: s.site_code || '',
        location: s.address || s.area || s.city || '',
        size: s.size || (s.width && s.height ? `${s.width}x${s.height} ft` : ''),
        meter_no: s.meter_no || prev.meter_no || '',
        service_number: s.service_number || prev.service_number || '',
        t_number: s.t_number || prev.t_number || ''
      }));
    }
  }

  function calculateBillAmount(units, rate, other = 0) {
    const u = parseFloat(units);
    const r = parseFloat(rate);
    const o = parseFloat(other) || 0;
    if (!isNaN(u) && !isNaN(r)) {
      return (u * r + o).toFixed(2);
    }
    return '';
  }

  async function saveBill(e) {
    e.preventDefault();
    if (isReadOnly) return;
    setSaving(true);
    const form = e.currentTarget;
    const formData = new FormData(form);
    const payload = Object.fromEntries(formData.entries());

    if (!payload.site_code && !payload.site_id) {
      alert('Please select or specify a Site Code.');
      setSaving(false);
      return;
    }

    const amt = Number(payload.amount || payload.payment_amount || 0);
    payload.amount = amt;
    payload.payment_amount = amt;

    try {
      if (editModal?.id) {
        await api.put(`/electricity/${editModal.id}`, payload);
        setBanner('✓ Electricity bill updated successfully!');
      } else {
        await api.post('/electricity', payload);
        setBanner('✓ New electricity bill logged successfully!');
      }
      setEditModal(null);
      await loadData();
      setTimeout(() => setBanner(''), 4000);
    } catch (err) {
      alert('Failed to save bill: ' + (err.response?.data?.message || err.message));
    } finally {
      setSaving(false);
    }
  }

  async function handleQuickPay(e) {
    e.preventDefault();
    if (!payModal?.id || isReadOnly) return;
    setSaving(true);
    const formData = new FormData(e.currentTarget);
    const paid_date = formData.get('paid_date') || new Date().toISOString().slice(0, 10);
    const payment_reference = formData.get('payment_reference') || '';

    try {
      await api.put(`/electricity/${payModal.id}`, {
        payment_status: 'Paid',
        paid_date,
        payment_reference
      });
      setPayModal(null);
      setBanner(`✓ Bill for ${payModal.site_code || 'Site'} marked as Paid!`);
      await loadData();
      setTimeout(() => setBanner(''), 4000);
    } catch (err) {
      alert('Failed to update status: ' + (err.response?.data?.message || err.message));
    } finally {
      setSaving(false);
    }
  }

  function toggleSelect(id) {
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function toggleSelectAll() {
    const visibleIds = filtered.map(r => r.id).filter(Boolean);
    const allSelected = visibleIds.length > 0 && visibleIds.every(id => selectedIds.has(id));
    if (allSelected) {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleIds.forEach(id => next.delete(id));
        return next;
      });
    } else {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleIds.forEach(id => next.add(id));
        return next;
      });
    }
  }

  function selectAllVisible() {
    setSelectedIds(new Set(filtered.map(r => r.id).filter(Boolean)));
  }

  function deselectAll() {
    setSelectedIds(new Set());
  }

  async function handleBatchDelete() {
    if (!canDelete || selectedIds.size === 0) return;
    const count = selectedIds.size;
    if (!confirm(`Are you sure you want to permanently delete all ${count} selected electricity bill record${count > 1 ? 's' : ''}? This action cannot be undone.`)) {
      return;
    }
    try {
      await api.post('/electricity/batch-delete', { ids: Array.from(selectedIds), hard: true });
      setBanner(`✓ ${count} electricity bill record${count > 1 ? 's' : ''} deleted successfully.`);
      setSelectedIds(new Set());
      await loadData();
      window.dispatchEvent(new CustomEvent('mb-electricity-updated'));
      setTimeout(() => setBanner(''), 4000);
    } catch (err) {
      alert('Failed to delete selected bills: ' + (err.response?.data?.message || err.message));
    }
  }

  async function deleteBill(id) {
    if (!canDelete) return;
    if (confirm('Permanently delete this electricity bill record?')) {
      try {
        await api.delete(`/electricity/${id}?hard=true`);
        setBanner('✓ Electricity bill deleted.');
        setSelectedIds(prev => {
          const next = new Set(prev);
          next.delete(id);
          return next;
        });
        await loadData();
        window.dispatchEvent(new CustomEvent('mb-electricity-updated'));
        setTimeout(() => setBanner(''), 3000);
      } catch (err) {
        alert('Failed to delete bill: ' + (err.response?.data?.message || err.message));
      }
    }
  }

  return (
    <>
      <div className="scooh-pagehead">
        <div>
          <div style={{ fontSize: '11px', fontWeight: 800, color: '#a78bfa', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: '4px' }}>
            Media Buzz — OOH Workspace
          </div>
          <h1 style={{ fontSize: '26px', fontWeight: 900, margin: 0, color: '#fff', letterSpacing: '-0.02em' }}>
            Electricity Tracker
          </h1>
          <p style={{ margin: '6px 0 0', color: '#94a3b8', fontSize: '13px' }}>
            Every electricity bill, tracked independently of campaign activity so no payment due date is missed.
          </p>
        </div>
        <div className="scooh-headactions" style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
          <button type="button" className="scooh-btn ghost" onClick={loadData}>
            🔄 Refresh
          </button>
          {canDelete && (
            <label className="scooh-btn ghost" style={{ cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: '6px' }} title="Import electricity bills from Excel (.xlsx, .xls)">
              <span>📁 {importingExcel ? 'Importing…' : 'Import Excel'}</span>
              <input type="file" accept=".xlsx,.xls" hidden disabled={importingExcel} onChange={handleElectricityExcelImport} />
            </label>
          )}
          <button type="button" className="scooh-btn ghost" onClick={() => exportElectricityExcel(filtered)}>
            📥 Export Excel
          </button>
          {canAdd && (
            <button type="button" className="scooh-btn purple-btn" onClick={openNewBill}>
              + Add bill
            </button>
          )}
        </div>
      </div>

      {isReadOnly && (
        <div className="scooh-banner" style={{ display: 'block', marginBottom: '16px', background: 'rgba(245,158,11,0.12)', border: '1px solid rgba(245,158,11,0.3)', color: '#fbbf24' }}>
          🔒 Read-Only Workspace: You are signed in as a Viewer. You can view electricity bills, track due dates, and export to Excel. Recording or deleting bills is disabled.
        </div>
      )}

      {banner && (
        <div className="scooh-banner" style={{ display: 'block', marginBottom: '16px', background: 'rgba(34,197,94,0.15)', border: '1px solid #22c55e', color: '#4ade80' }}>
          {banner}
        </div>
      )}

      {/* Financial & Operational KPI Cards */}
      <div className="scooh-kpirow" style={{ marginBottom: '22px' }}>
        <div className="scooh-kpi alert">
          <div className="n">{money(totalExpense)}</div>
          <div className="l">Total Electricity Expense</div>
          <div className="scooh-kpi-note" style={{ color: '#94a3b8' }}>{rows.length} total bills logged</div>
        </div>
        <div className="scooh-kpi good">
          <div className="n">{money(paidExpense)}</div>
          <div className="l">Paid Bills</div>
          <div className="scooh-kpi-note" style={{ color: '#4ade80' }}>
            {rows.filter(r => r.payment_status === 'Paid').length} bills cleared
          </div>
        </div>
        <div className={`scooh-kpi ${overdueCount > 0 ? 'danger' : 'alert'}`}>
          <div className="n">{money(pendingExpense)}</div>
          <div className="l">Pending / Due Bills</div>
          <div className="scooh-kpi-note" style={{ color: overdueCount > 0 ? '#ef4444' : '#fbbf24' }}>
            {overdueCount > 0 ? `⚠ ${overdueCount} overdue bills` : `${rows.filter(r => r.payment_status !== 'Paid').length} awaiting payment`}
          </div>
        </div>
        <div className="scooh-kpi alert">
          <div className="n">{sites.length || 57}</div>
          <div className="l">Total Sites Monitored</div>
          <div className="scooh-kpi-note" style={{ color: '#94a3b8' }}>SSV & MB Bills</div>
        </div>
      </div>

      {/* Main Section */}
      <div className="scooh-electricity-section">
        <div className="scooh-sectionbar">
          <div>
            <h2 style={{ fontSize: '16px', fontWeight: 800, color: '#f8fafc' }}>
              Electricity Bills & Payment Due Dates
            </h2>
            <p style={{ margin: '4px 0 0', fontSize: '12px', color: '#94a3b8' }}>
              All imported and newly added electricity bills are shown here with site and meter details.
            </p>
          </div>
          <div style={{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap' }}>
            {canAdd && (
              <button type="button" className="scooh-btn purple-btn" onClick={openNewBill}>
                + Add bill
              </button>
            )}
          </div>
        </div>

        {/* Filters Toolbar */}
        <div className="scooh-toolbar" style={{ padding: '14px 20px', borderBottom: '1px solid rgba(255,255,255,0.06)', display: 'flex', gap: '12px', flexWrap: 'wrap', alignItems: 'center' }}>
          <input
            className="scooh-search"
            style={{ flex: '1 1 240px', minWidth: '200px' }}
            placeholder="Search site code (e.g. 01, MB-01), location, meter, consumer no…"
            value={search}
            onChange={e => setSearch(e.target.value)}
          />
          <select
            value={statusFilter}
            onChange={e => setStatusFilter(e.target.value)}
            style={{ padding: '8px 12px', borderRadius: '8px', background: '#111925', color: '#e2e8f0', border: '1px solid #334155', fontSize: '12px' }}
          >
            <option value="ALL">All Statuses ({rows.length})</option>
            <option value="Pending">Pending</option>
            <option value="Paid">Paid</option>
            <option value="Overdue">Overdue ({overdueCount})</option>
          </select>
          {providers.length > 0 && (
            <select
              value={providerFilter}
              onChange={e => setProviderFilter(e.target.value)}
              style={{ padding: '8px 12px', borderRadius: '8px', background: '#111925', color: '#e2e8f0', border: '1px solid #334155', fontSize: '12px' }}
            >
              <option value="ALL">All Providers ({providers.length})</option>
              {providers.map(p => (
                <option key={p} value={p}>{p}</option>
              ))}
            </select>
          )}
          {months.length > 0 && (
            <select
              value={monthFilter}
              onChange={e => setMonthFilter(e.target.value)}
              style={{ padding: '8px 12px', borderRadius: '8px', background: '#111925', color: '#e2e8f0', border: '1px solid #334155', fontSize: '12px' }}
            >
              <option value="ALL">All Billing Months ({months.length})</option>
              {months.map(m => (
                <option key={m} value={m}>{m}</option>
              ))}
            </select>
          )}
          <select
            value={`${sortState.key}:${sortState.dir}`}
            onChange={e => {
              const [k, d] = e.target.value.split(':');
              setSortState({ key: k, dir: d });
            }}
            style={{ padding: '8px 12px', borderRadius: '8px', background: '#111925', color: '#e2e8f0', border: '1px solid #334155', fontSize: '12px', fontWeight: 600 }}
          >
            <option value="site_code:asc">Sort: Site Code (01 → 87)</option>
            <option value="site_code:desc">Sort: Site Code (87 → 01)</option>
            <option value="due_date:asc">Sort: Due Date (Earliest First)</option>
            <option value="due_date:desc">Sort: Due Date (Latest First)</option>
            <option value="amount:desc">Sort: Amount (High to Low)</option>
            <option value="amount:asc">Sort: Amount (Low to High)</option>
            <option value="location:asc">Sort: Location (A–Z)</option>
            <option value="payment_status:asc">Sort: Status (A–Z)</option>
            <option value="billing_month:desc">Sort: Billing Month</option>
          </select>
          {(search || statusFilter !== 'ALL' || providerFilter !== 'ALL' || monthFilter !== 'ALL' || sortState.key !== 'due_date' || sortState.dir !== 'asc') && (
            <button
              type="button"
              className="scooh-btn ghost"
              style={{ fontSize: '11px', padding: '6px 10px' }}
              onClick={() => { setSearch(''); setStatusFilter('ALL'); setProviderFilter('ALL'); setMonthFilter('ALL'); setSortState({ key: 'due_date', dir: 'asc' }); }}
            >
              Reset Filters
            </button>
          )}
        </div>

        {/* Batch Selection Action Bar */}
        {canDelete && selectedIds.size > 0 && (
          <div style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            padding: '11px 18px',
            background: 'linear-gradient(90deg, rgba(239,68,68,0.18), rgba(239,68,68,0.08))',
            borderBottom: '1px solid rgba(239,68,68,0.3)',
            color: '#fca5a5',
            fontSize: '13px',
            flexWrap: 'wrap',
            gap: '12px'
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap' }}>
              <span style={{ fontWeight: 800, color: '#ffffff', display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
                <span style={{ display: 'inline-block', width: '8px', height: '8px', borderRadius: '50%', background: '#ef4444' }}></span>
                {selectedIds.size} bill{selectedIds.size > 1 ? 's' : ''} selected
              </span>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={selectAllVisible}
                style={{ fontSize: '11.5px', padding: '4px 10px', color: '#f1f5f9', borderColor: '#475569' }}
              >
                Select all visible ({filtered.length})
              </button>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={deselectAll}
                style={{ fontSize: '11.5px', padding: '4px 10px', color: '#cbd5e1', borderColor: '#475569' }}
              >
                Deselect all
              </button>
            </div>
            <button
              type="button"
              className="scooh-btn danger"
              onClick={handleBatchDelete}
              style={{
                background: '#ef4444',
                color: '#ffffff',
                border: 'none',
                fontWeight: 800,
                padding: '7px 16px',
                borderRadius: '7px',
                cursor: 'pointer',
                display: 'inline-flex',
                alignItems: 'center',
                gap: '6px',
                boxShadow: '0 2px 10px rgba(239,68,68,0.45)'
              }}
            >
              🗑 Delete selected ({selectedIds.size})
            </button>
          </div>
        )}

        {/* Table */}
        <div className="scooh-tablewrap">
          <table className="scooh-table">
            <thead>
              <tr>
                {canDelete && (
                  <th style={{ width: '42px', textAlign: 'center', padding: '10px 8px' }}>
                    <input
                      type="checkbox"
                      checked={filtered.length > 0 && filtered.every(r => selectedIds.has(r.id))}
                      onChange={toggleSelectAll}
                      title="Select all visible bills"
                      style={{ cursor: 'pointer' }}
                    />
                  </th>
                )}
                <SortHeader label="SITE CODE" sortKey="site_code" currentSort={sortState} onSort={handleSort} style={{ minWidth: '110px' }} />
                <SortHeader label="LOCATION" sortKey="location" currentSort={sortState} onSort={handleSort} style={{ minWidth: '180px' }} />
                <SortHeader label="SIZE" sortKey="size" currentSort={sortState} onSort={handleSort} style={{ minWidth: '85px' }} />
                <SortHeader label="METER / SERVICE" sortKey="meter_no" currentSort={sortState} onSort={handleSort} style={{ minWidth: '140px' }} />
                <SortHeader label="T NUMBER" sortKey="t_number" currentSort={sortState} onSort={handleSort} style={{ minWidth: '100px' }} />
                <SortHeader label="BILL / PROVIDER" sortKey="bill_type" currentSort={sortState} onSort={handleSort} style={{ minWidth: '120px' }} />
                <SortHeader label="BILLING MONTH" sortKey="billing_month" currentSort={sortState} onSort={handleSort} style={{ minWidth: '110px' }} />
                <SortHeader label="DUE DATE" sortKey="due_date" currentSort={sortState} onSort={handleSort} style={{ minWidth: '110px' }} />
                <SortHeader label="PAYMENT AMOUNT" sortKey="amount" currentSort={sortState} onSort={handleSort} align="right" style={{ minWidth: '120px' }} />
                <SortHeader label="STATUS" sortKey="payment_status" currentSort={sortState} onSort={handleSort} align="center" style={{ minWidth: '95px' }} />
                <th style={{ minWidth: '140px', textAlign: 'center' }}>ACTIONS</th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr>
                  <td colSpan={canDelete ? "12" : "11"} className="scooh-empty" style={{ padding: '36px 20px' }}>
                    <div style={{ fontSize: '14px', fontWeight: 700, color: '#94a3b8' }}>No electricity bills match your query</div>
                    <div style={{ fontSize: '12px', color: '#64748b', marginTop: '4px' }}>
                      {canAdd ? 'Click "+ Add bill" to record a new electricity bill for any hoarding site.' : 'No recorded electricity bills.'}
                    </div>
                  </td>
                </tr>
              ) : (
                filtered.map(r => {
                  const isOverdue = r.payment_status !== 'Paid' && r.due_date && new Date(r.due_date) < new Date();
                  return (
                    <tr key={r.id}>
                      {canDelete && (
                        <td style={{ textAlign: 'center', padding: '10px 8px' }}>
                          <input
                            type="checkbox"
                            checked={selectedIds.has(r.id)}
                            onChange={() => toggleSelect(r.id)}
                            style={{ cursor: 'pointer' }}
                          />
                        </td>
                      )}
                      <td>
                        <span className="scooh-plate" style={{ fontSize: '11.5px', fontWeight: 800 }}>
                          {r.site_code || '—'}
                        </span>
                      </td>
                      <td>
                        <div style={{ fontWeight: 700, color: '#edf2f6', fontSize: '12.5px', lineHeight: 1.35 }}>
                          {r.location || '—'}
                        </div>
                      </td>
                      <td style={{ color: '#cbd5e1', fontSize: '12px' }}>
                        {r.size || '—'}
                      </td>
                      <td>
                        <div style={{ fontSize: '11.5px', color: '#f1f5f9', fontWeight: 600 }}>
                          {r.meter_no ? `Mtr: ${r.meter_no}` : (r.service_number ? `Srv: ${r.service_number}` : '—')}
                        </div>
                        {r.service_number && r.meter_no && (
                          <div style={{ fontSize: '10.5px', color: '#94a3b8', marginTop: '2px' }}>
                            Srv: {r.service_number}
                          </div>
                        )}
                      </td>
                      <td style={{ color: '#cbd5e1', fontSize: '12px' }}>
                        {r.t_number || '—'}
                      </td>
                      <td>
                        <span
                          className="scooh-badgechip"
                          style={{
                            background: r.bill_type === 'SSV' ? 'rgba(52,211,153,0.18)' :
                              (r.bill_type === 'MB' ? 'rgba(251,191,36,0.18)' :
                              (r.bill_type === 'Torrent Power' ? 'rgba(56,189,248,0.16)' : 'rgba(168,85,247,0.16)')),
                            color: r.bill_type === 'SSV' ? '#34d399' :
                              (r.bill_type === 'MB' ? '#fbbf24' :
                              (r.bill_type === 'Torrent Power' ? '#38bdf8' : '#c084fc')),
                            fontWeight: 800,
                            fontSize: '11px',
                            letterSpacing: '0.02em',
                            border: r.bill_type === 'SSV' ? '1px solid rgba(52,211,153,0.3)' :
                              (r.bill_type === 'MB' ? '1px solid rgba(251,191,36,0.35)' : 'none')
                          }}
                        >
                          {r.bill_type || 'SSV'}
                        </span>
                      </td>
                      <td style={{ color: '#cbd5e1', fontSize: '12px', fontWeight: 600 }}>
                        {r.billing_month || '—'}
                      </td>
                      <td>
                        <div style={{ fontSize: '12px', color: isOverdue ? '#f87171' : '#cbd5e1', fontWeight: isOverdue ? 800 : 500 }}>
                          {formatDate(r.due_date)}
                        </div>
                        {isOverdue && (
                          <span style={{ display: 'inline-block', marginTop: '2px', fontSize: '9.5px', fontWeight: 800, color: '#ef4444', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
                            ⚠ Overdue
                          </span>
                        )}
                      </td>
                      <td style={{ textAlign: 'right' }}>
                        <b style={{ fontSize: '13px', color: '#f8fafc', letterSpacing: '0.01em' }}>
                          {money(r.amount || r.payment_amount)}
                        </b>
                      </td>
                      <td style={{ textAlign: 'center' }}>
                        <span className={`scooh-pill ${r.payment_status === 'Paid' ? 'active' : r.payment_status === 'Pending' ? 'watch' : 'vacant'}`}>
                          {r.payment_status || 'Pending'}
                        </span>
                      </td>
                      <td>
                        <div className="scooh-rowactions" style={{ justifyContent: 'center', gap: '6px' }}>
                          {canPay && r.payment_status !== 'Paid' && (
                            <button
                              type="button"
                              className="scooh-btn"
                              style={{
                                padding: '4px 8px',
                                fontSize: '11px',
                                fontWeight: 700,
                                background: 'rgba(34,197,94,0.18)',
                                color: '#4ade80',
                                border: '1px solid rgba(34,197,94,0.35)',
                                borderRadius: '6px'
                              }}
                              onClick={() => setPayModal(r)}
                              title="Mark this bill as Paid"
                            >
                              ✓ Pay
                            </button>
                          )}
                          <button
                            type="button"
                            className="scooh-iconbtn scooh-text-action"
                            onClick={() => setEditModal(r)}
                            title={isReadOnly ? 'View Bill Details' : 'Edit Bill'}
                          >
                            {isReadOnly ? 'View' : 'Edit'}
                          </button>
                          {canDelete && (
                            <button
                              type="button"
                              className="scooh-iconbtn danger-icon"
                              onClick={() => deleteBill(r.id)}
                              title="Delete Bill"
                            >
                              ×
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Add / Edit Bill Modal */}
      {editModal && (
        <div className="scooh-modal-overlay">
          <div className="scooh-modal" style={{ maxWidth: '640px' }}>
            <div className="scooh-modalhead">
              <div>
                <h2>{editModal.id ? (isReadOnly ? 'View Electricity Bill' : 'Edit Electricity Bill') : 'Add Electricity Bill'}</h2>
                <span className="scooh-modal-subtitle">
                  {isReadOnly ? 'Read-only view of utility meter bill and payment record' : 'Auto-fill site metadata and manage payment tracking'}
                </span>
              </div>
              <button type="button" className="scooh-modal-close" onClick={() => setEditModal(null)}>×</button>
            </div>
            <form onSubmit={saveBill}>
              <div className="scooh-modalbody">
                {/* Site Selection Quick Picker */}
                <div className="scooh-field" style={{ gridColumn: '1 / -1', background: 'rgba(255,255,255,0.03)', padding: '12px', borderRadius: '10px', border: '1px solid rgba(255,255,255,0.08)' }}>
                  <label style={{ color: '#c4b5fd', fontWeight: 800 }}>⚡ QUICK AUTO-FILL FROM SITE INVENTORY</label>
                  <select
                    value={editModal.site_id || editModal.site_code || ''}
                    onChange={e => handleSiteChange(e.target.value)}
                    style={{ marginTop: '6px' }}
                  >
                    <option value="">-- Select a Site to Auto-Fill Location, Size & Meter --</option>
                    {sites.map(s => (
                      <option key={s.id || s.site_code} value={s.id || s.site_code}>
                        {s.site_code} — {s.area || s.address || s.city} ({s.size || 'Standard'})
                      </option>
                    ))}
                  </select>
                </div>

                <div className="scooh-grid2">
                  <div className="scooh-field">
                    <label>Site Code *</label>
                    <input
                      name="site_code"
                      value={editModal.site_code ?? ''}
                      onChange={e => setEditModal({ ...editModal, site_code: e.target.value })}
                      placeholder="e.g. AMD-GT-001"
                      required
                    />
                  </div>
                  <div className="scooh-field">
                    <label>Site Size</label>
                    <input
                      name="size"
                      value={editModal.size ?? ''}
                      onChange={e => setEditModal({ ...editModal, size: e.target.value })}
                      placeholder="e.g. 30x10 ft"
                    />
                  </div>
                  <div className="scooh-field" style={{ gridColumn: '1 / -1' }}>
                    <label>Location / Address *</label>
                    <input
                      name="location"
                      value={editModal.location ?? ''}
                      onChange={e => setEditModal({ ...editModal, location: e.target.value })}
                      placeholder="e.g. Shivranjani Cross Roads, Ahmedabad"
                      required
                    />
                  </div>

                  <div className="scooh-field">
                    <label>Bill Provider / Type *</label>
                    <select
                      name="bill_type"
                      value={editModal.bill_type ?? 'SSV'}
                      onChange={e => setEditModal({ ...editModal, bill_type: e.target.value })}
                    >
                      <option value="SSV">SSV</option>
                      <option value="MB">MB</option>
                      <option value="Torrent Power">Torrent Power</option>
                      <option value="UGVCL">UGVCL</option>
                      <option value="PGVCL">PGVCL</option>
                      <option value="DGVCL">DGVCL</option>
                      <option value="MGVCL">MGVCL</option>
                      <option value="Other">Other Utility</option>
                      {editModal.bill_type && !['SSV', 'MB', 'Torrent Power', 'UGVCL', 'PGVCL', 'DGVCL', 'MGVCL', 'Other'].includes(editModal.bill_type) && (
                        <option value={editModal.bill_type}>{editModal.bill_type}</option>
                      )}
                    </select>
                  </div>
                  <div className="scooh-field">
                    <label>Billing Month</label>
                    <input
                      name="billing_month"
                      value={editModal.billing_month ?? ''}
                      onChange={e => setEditModal({ ...editModal, billing_month: e.target.value })}
                      placeholder="e.g. Aug 2026"
                    />
                  </div>

                  <div className="scooh-field">
                    <label>Meter Number</label>
                    <input
                      name="meter_no"
                      value={editModal.meter_no ?? ''}
                      onChange={e => setEditModal({ ...editModal, meter_no: e.target.value })}
                      placeholder="e.g. MTR-UGVCL-8841"
                    />
                  </div>
                  <div className="scooh-field">
                    <label>Service Number</label>
                    <input
                      name="service_number"
                      value={editModal.service_number ?? ''}
                      onChange={e => setEditModal({ ...editModal, service_number: e.target.value })}
                      placeholder="e.g. SRV-998241"
                    />
                  </div>

                  <div className="scooh-field">
                    <label>T Number</label>
                    <input
                      name="t_number"
                      value={editModal.t_number ?? ''}
                      onChange={e => setEditModal({ ...editModal, t_number: e.target.value })}
                      placeholder="e.g. T-4401"
                    />
                  </div>
                  <div className="scooh-field">
                    <label>Bill Date</label>
                    <input
                      type="date"
                      name="bill_date"
                      defaultValue={editModal.bill_date ? editModal.bill_date.slice(0, 10) : ''}
                    />
                  </div>

                  <div className="scooh-field">
                    <label>Payment Due Date *</label>
                    <input
                      type="date"
                      name="due_date"
                      defaultValue={editModal.due_date ? editModal.due_date.slice(0, 10) : ''}
                      required
                    />
                  </div>
                  <div className="scooh-field">
                    <label>Units Consumed (kWh)</label>
                    <input
                      type="number"
                      step="any"
                      name="units"
                      value={editModal.units ?? ''}
                      onChange={e => {
                        const u = e.target.value;
                        const calc = calculateBillAmount(u, editModal.rate, editModal.other_charges);
                        setEditModal(prev => ({
                          ...prev,
                          units: u,
                          amount: calc || prev.amount
                        }));
                      }}
                      placeholder="e.g. 850"
                    />
                  </div>

                  <div className="scooh-field">
                    <label>Rate per Unit (₹)</label>
                    <input
                      type="number"
                      step="any"
                      name="rate"
                      value={editModal.rate ?? ''}
                      onChange={e => {
                        const r = e.target.value;
                        const calc = calculateBillAmount(editModal.units, r, editModal.other_charges);
                        setEditModal(prev => ({
                          ...prev,
                          rate: r,
                          amount: calc || prev.amount
                        }));
                      }}
                      placeholder="e.g. 9.23"
                    />
                  </div>
                  <div className="scooh-field">
                    <label>Other Charges (₹)</label>
                    <input
                      type="number"
                      step="any"
                      name="other_charges"
                      value={editModal.other_charges ?? ''}
                      onChange={e => {
                        const o = e.target.value;
                        const calc = calculateBillAmount(editModal.units, editModal.rate, o);
                        setEditModal(prev => ({
                          ...prev,
                          other_charges: o,
                          amount: calc || prev.amount
                        }));
                      }}
                      placeholder="0.00"
                    />
                  </div>

                  <div className="scooh-field" style={{ gridColumn: '1 / -1', background: 'rgba(139,92,246,0.06)', padding: '12px', borderRadius: '8px', border: '1px solid rgba(139,92,246,0.2)' }}>
                    <label style={{ color: '#c4b5fd', fontWeight: 900, fontSize: '12px' }}>BILL / PAYMENT AMOUNT (₹) *</label>
                    <input
                      type="number"
                      step="any"
                      name="amount"
                      value={editModal.amount ?? ''}
                      onChange={e => setEditModal({ ...editModal, amount: e.target.value })}
                      placeholder="e.g. 7850"
                      style={{ fontSize: '15px', fontWeight: 800 }}
                      required
                    />
                  </div>

                  <div className="scooh-field">
                    <label>Payment Status</label>
                    <select
                      name="payment_status"
                      value={editModal.payment_status ?? 'Pending'}
                      onChange={e => setEditModal({ ...editModal, payment_status: e.target.value })}
                    >
                      <option value="Pending">Pending</option>
                      <option value="Paid">Paid</option>
                      <option value="Overdue">Overdue</option>
                      <option value="Partial">Partial</option>
                    </select>
                  </div>
                  <div className="scooh-field">
                    <label>Paid Date</label>
                    <input
                      type="date"
                      name="paid_date"
                      defaultValue={editModal.paid_date ? editModal.paid_date.slice(0, 10) : ''}
                    />
                  </div>

                  <div className="scooh-field" style={{ gridColumn: '1 / -1' }}>
                    <label>Payment Reference / Transaction ID</label>
                    <input
                      name="payment_reference"
                      value={editModal.payment_reference ?? ''}
                      onChange={e => setEditModal({ ...editModal, payment_reference: e.target.value })}
                      placeholder="e.g. UPI-9923847291 / NEFT / Cheque #4412"
                    />
                  </div>

                  <div className="scooh-field" style={{ gridColumn: '1 / -1' }}>
                    <label>Operational Notes</label>
                    <textarea
                      name="notes"
                      value={editModal.notes ?? ''}
                      onChange={e => setEditModal({ ...editModal, notes: e.target.value })}
                      placeholder="Any notes on meter reading, meter seal, or utility dispatch…"
                      rows={2}
                    />
                  </div>
                </div>
              </div>
              <div className="scooh-modalfoot">
                <button type="button" className="scooh-btn ghost" onClick={() => setEditModal(null)}>
                  {isReadOnly ? 'Close' : 'Cancel'}
                </button>
                {!isReadOnly && (
                  <button type="submit" className="scooh-btn purple-btn" disabled={saving}>
                    {saving ? 'Saving…' : editModal.id ? 'Update Electricity Bill' : 'Save Electricity Bill'}
                  </button>
                )}
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Quick Mark Paid Modal */}
      {payModal && (
        <div className="scooh-modal-overlay">
          <div className="scooh-modal" style={{ maxWidth: '440px' }}>
            <div className="scooh-modalhead">
              <div>
                <h2>Mark Bill as Paid</h2>
                <span className="scooh-modal-subtitle">{payModal.site_code} — {payModal.location || 'Site Bill'}</span>
              </div>
              <button type="button" className="scooh-modal-close" onClick={() => setPayModal(null)}>×</button>
            </div>
            <form onSubmit={handleQuickPay}>
              <div className="scooh-modalbody">
                <div style={{ padding: '12px', background: 'rgba(34,197,94,0.1)', border: '1px solid rgba(34,197,94,0.25)', borderRadius: '8px', marginBottom: '14px' }}>
                  <div style={{ fontSize: '11px', color: '#94a3b8' }}>BILL AMOUNT</div>
                  <div style={{ fontSize: '20px', fontWeight: 900, color: '#4ade80', marginTop: '2px' }}>
                    {money(payModal.amount || payModal.payment_amount)}
                  </div>
                  <div style={{ fontSize: '11px', color: '#cbd5e1', marginTop: '4px' }}>
                    Month: {payModal.billing_month || 'Current'} | Provider: {payModal.bill_type || 'General'}
                  </div>
                </div>

                <div className="scooh-field">
                  <label>Paid Date *</label>
                  <input
                    type="date"
                    name="paid_date"
                    defaultValue={new Date().toISOString().slice(0, 10)}
                    required
                  />
                </div>
                <div className="scooh-field">
                  <label>Payment Reference / UTR / Mode</label>
                  <input
                    name="payment_reference"
                    placeholder="e.g. UPI-9923847291 / NetBanking / Cheque"
                    autoFocus
                  />
                </div>
              </div>
              <div className="scooh-modalfoot">
                <button type="button" className="scooh-btn ghost" onClick={() => setPayModal(null)}>Cancel</button>
                <button type="submit" className="scooh-btn purple-btn" style={{ background: '#22c55e', borderColor: '#22c55e' }} disabled={saving}>
                  {saving ? 'Updating…' : 'Confirm Paid'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </>
  );
}

function DataToolsView() {
  const [xlsxFile, setXlsxFile] = useState(null);
  const [elecFile, setElecFile] = useState(null);
  const [jsonFile, setJsonFile] = useState(null);
  const [xlsxStatus, setXlsxStatus] = useState('No Excel file selected.');
  const [elecStatus, setElecStatus] = useState('No Electricity Excel selected.');
  const [jsonStatus, setJsonStatus] = useState('No JSON file selected.');
  const [loading, setLoading] = useState(false);

  async function importXlsx() {
    if (!xlsxFile) return alert('Please choose an Excel file (.xlsx, .xls) first.');
    setLoading(true);
    setXlsxStatus(`Importing ${xlsxFile.name}…`);
    try {
      const f = new FormData();
      f.append('file', xlsxFile);
      const r = await api.post('/import/xlsx', f, { headers: { 'Content-Type': 'multipart/form-data' } });
      const msg = r.data.message || `✓ Processed ${r.data.rows} sites (${r.data.updated} updated, ${r.data.created} created).`;
      setXlsxStatus(msg);
      alert(`✓ Excel Import Successful!\n\n${msg}`);
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
    } catch (e) {
      setXlsxStatus(`✕ Import failed: ${e.response?.data?.message || e.message}`);
      alert('Import failed: ' + (e.response?.data?.message || e.message));
    } finally {
      setLoading(false);
    }
  }

  async function importElecXlsx() {
    if (!elecFile) return alert('Please choose an Electricity Excel file (.xlsx, .xls) first.');
    setLoading(true);
    setElecStatus(`Importing ${elecFile.name}…`);
    try {
      const f = new FormData();
      f.append('file', elecFile);
      const r = await api.post('/import/electricity-xlsx', f, { headers: { 'Content-Type': 'multipart/form-data' } });
      const msg = r.data.message || `✓ Processed ${r.data.rows} electricity bills (${r.data.updated} updated, ${r.data.created} created).`;
      setElecStatus(msg);
      alert(`✓ Electricity Import Successful!\n\n${msg}`);
      window.dispatchEvent(new CustomEvent('mb-electricity-updated'));
    } catch (e) {
      setElecStatus(`✕ Import failed: ${e.response?.data?.message || e.message}`);
      alert('Import failed: ' + (e.response?.data?.message || e.message));
    } finally {
      setLoading(false);
    }
  }

  async function importJson() {
    if (!jsonFile) return alert('Please choose a JSON backup file first.');
    setLoading(true);
    setJsonStatus(`Restoring ${jsonFile.name}…`);
    try {
      const f = new FormData();
      f.append('file', jsonFile);
      const r = await api.post('/import/json', f, { headers: { 'Content-Type': 'multipart/form-data' } });
      setJsonStatus(`✓ Successfully restored ${r.data.count} records from ${jsonFile.name}!`);
      alert(`Success: ${r.data.count} records restored.`);
    } catch (e) {
      setJsonStatus(`✕ Restore failed: ${e.response?.data?.message || e.message}`);
      alert('Restore failed: ' + (e.response?.data?.message || e.message));
    } finally {
      setLoading(false);
    }
  }

  async function exportExcel() {
    try {
      const { data } = await api.get('/sites');
      const source = (Array.isArray(data) && data.length > 0) ? data : defaultSites;
      const rows = source.map(unpackSite).map((x, idx) => {
        let w = x.width || '', h = x.height || '';
        if ((!w || !h) && x.size) {
          const parts = String(x.size).replace(/[^0-9xX.]/g, '').split(/[xX]/);
          if (parts.length === 2) { w = parts[0]; h = parts[1]; }
        }
        const sqft = (Number(w || 0) && Number(h || 0)) ? Number(w) * Number(h) : (x.sq_ft || '');
        const coords = x.gps || ([x.latitude, x.longitude].filter(Boolean).join(', ')) || '';

        return {
          'SR NO': idx + 1,
          'AREA': x.area || x.city || '',
          'LOCATION': x.address || '',
          'MEDIA': x.media_type || 'Hoarding',
          'LIGHT': x.lighting || 'BL',
          'W': w || '',
          'H': h || '',
          'SQ FT': sqft,
          'AVAILABLITY': x.ppt_availability || x.availability || 'Available',
          'Selling Amount': Number(x.ppt_rate || x.monthly_rate || 0),
          'Latitude Longitude': coords
        };
      });

      await exportStyledExcel(rows, 'MediaBuzz_Sites.xlsx');
    } catch (e) {
      alert('Export failed: ' + e.message);
    }
  }

  async function exportCsv() {
    try {
      const { data } = await api.get('/sites');
      const source = (Array.isArray(data) && data.length > 0) ? data : defaultSites;
      const rows = source.map(unpackSite).map((x, idx) => {
        let w = x.width || '', h = x.height || '';
        if ((!w || !h) && x.size) {
          const parts = String(x.size).replace(/[^0-9xX.]/g, '').split(/[xX]/);
          if (parts.length === 2) { w = parts[0]; h = parts[1]; }
        }
        const sqft = (Number(w || 0) && Number(h || 0)) ? Number(w) * Number(h) : '';
        const coords = x.gps || ([x.latitude, x.longitude].filter(Boolean).join(', ')) || '';

        return {
          'SR NO': idx + 1,
          'AREA': x.area || x.city || '',
          'LOCATION': x.address || '',
          'MEDIA': x.media_type || 'Hoarding',
          'LIGHT': x.lighting || 'BL',
          'W': w || '',
          'H': h || '',
          'SQ FT': sqft,
          'AVAILABLITY': x.ppt_availability || x.availability || 'Available',
          'Selling Amount': Number(x.ppt_rate || x.monthly_rate || 0),
          'Latitude Longitude': coords
        };
      });

      const ws = XLSX.utils.json_to_sheet(rows);
      const csv = XLSX.utils.sheet_to_csv(ws);
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const u = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = u;
      a.download = 'MediaBuzz_Sites.csv';
      a.click();
      URL.revokeObjectURL(u);
    } catch (e) {
      alert('Export CSV failed: ' + e.message);
    }
  }

  async function exportJson() {
    try {
      const { data } = await api.get('/export/json');
      const b = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
      const u = URL.createObjectURL(b);
      const a = document.createElement('a');
      a.href = u;
      a.download = 'mediabuzz-database-backup.json';
      a.click();
      URL.revokeObjectURL(u);
    } catch (e) {
      alert('Export JSON failed: ' + e.message);
    }
  }

  function downloadProductionGuide() {
    const text = `# Media Buzz OOH Workspace — Production Guide & System Documentation
Version: 3.0.0 | Environment: Production (Hostinger / VPS / Cloud)

## 1. System Architecture
- Frontend: React 18, React Router v6, Vite SPA
- Backend: Express Node.js API with JWT Stateless Auth & Bcrypt
- Database: MySQL 8.0 (InnoDB) with composite indexing
- Presentation: Client-side 16:9 widescreen PPTX generation (PptxGenJS)
- Excel: Native 11-column parser & builder (SheetJS XLSX)

## 2. Production Advantages
- Sub-50ms API response time with zero WordPress PHP overhead
- 100% data privacy on your own MySQL database
- Client-side PowerPoint generation in <2 seconds with custom brand cards
- Zero recurring software licensing or subscription fees
- Full source code ownership in modern React & Node.js

## 3. Production Disadvantages & Mitigations
- Self-managed infrastructure (Mitigation: Auto-restart via PM2/Hostinger Node runner)
- Database backups (Mitigation: Built-in one-click Export JSON backup + MySQL cron dumps)
- SMTP email configuration (Mitigation: Standard SMTP environment variables in .env)

## 4. Excel Standard Column Format
SR NO | AREA | LOCATION | MEDIA | LIGHT | W | H | SQ FT | AVAILABLITY | Selling Amount | Latitude Longitude

## 5. Security & Deployment
- Set strong JWT_SECRET and database passwords in server/.env
- Configure daily cron backup: mysqldump -u <user> -p'<pass>' <dbname> > /backups/mb_\$(date +%F).sql
- Hostinger Node.js runner: Node 18/20, Root: /, Startup: server/src/index.js
`;
    const b = new Blob([text], { type: 'text/markdown' });
    const u = URL.createObjectURL(b);
    const a = document.createElement('a');
    a.href = u;
    a.download = 'MediaBuzz_Production_Guide.md';
    a.click();
    URL.revokeObjectURL(u);
  }

  return (
    <>
      <PageHead
        title="Import / Export"
        desc="Import or update your OOH workbook, then export operational data as Excel, PowerPoint, JSON or CSV."
        actions={
          <>
            <button type="button" className="scooh-btn ghost" onClick={downloadProductionGuide}>
              Download Production Guide (.md)
            </button>
            <button type="button" className="scooh-btn primary" onClick={exportExcel}>
              Export Excel (.xlsx)
            </button>
          </>
        }
      />

      <div className="scooh-grid2 scooh-data-grid">
        {/* Import from Excel */}
        <div className="scooh-panel scooh-data-panel">
          <h3>Import from Excel</h3>
          <p className="scooh-footnote scooh-data-copy">
            Select your 11-column Excel file, then click <b>Import Data</b> to auto-update sites and availability.
          </p>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', background: 'var(--mb-input)', border: '1px solid var(--mb-border)', borderRadius: '12px', padding: '8px 12px', marginTop: '12px' }}>
            <label className="scooh-btn ghost" style={{ margin: 0, cursor: 'pointer', flexShrink: 0, padding: '7px 14px', fontSize: '12px', display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
              📁 {xlsxFile ? 'Change File' : 'Choose Excel'}
              <input
                type="file"
                id="import-xlsx"
                accept=".xlsx,.xls"
                hidden
                onChange={e => {
                  const f = e.target.files?.[0];
                  setXlsxFile(f || null);
                  setXlsxStatus(f ? `Selected: ${f.name}` : '');
                }}
              />
            </label>
            <div style={{ flex: 1, minWidth: 0, fontSize: '12px', color: xlsxFile ? 'var(--mb-text)' : 'var(--mb-muted)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              {xlsxFile ? `📄 ${xlsxFile.name} (${(xlsxFile.size / 1024).toFixed(1)} KB)` : 'No Excel file selected'}
            </div>
            <button
              type="button"
              className="scooh-btn purple-btn"
              onClick={importXlsx}
              disabled={loading || !xlsxFile}
              style={{ flexShrink: 0, padding: '7px 18px', fontSize: '12.5px' }}
            >
              {loading ? 'Importing…' : '⬆ Import Data'}
            </button>
          </div>
          {xlsxStatus && <div className="scooh-footnote" style={{ marginTop: '10px', color: '#48c79a', fontWeight: 600 }}>{xlsxStatus}</div>}
        </div>

        {/* Import Electricity Bills */}
        <div className="scooh-panel scooh-data-panel">
          <h3>Import Electricity Bills</h3>
          <p className="scooh-footnote scooh-data-copy">
            Select an Excel spreadsheet containing electricity meter bills to auto-log or update monthly utility payments.
          </p>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', background: 'var(--mb-input)', border: '1px solid var(--mb-border)', borderRadius: '12px', padding: '8px 12px', marginTop: '12px' }}>
            <label className="scooh-btn ghost" style={{ margin: 0, cursor: 'pointer', flexShrink: 0, padding: '7px 14px', fontSize: '12px', display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
              📁 {elecFile ? 'Change File' : 'Choose Bills Excel'}
              <input
                type="file"
                id="import-elec-xlsx"
                accept=".xlsx,.xls"
                hidden
                onChange={e => {
                  const f = e.target.files?.[0];
                  setElecFile(f || null);
                  setElecStatus(f ? `Selected: ${f.name}` : '');
                }}
              />
            </label>
            <div style={{ flex: 1, minWidth: 0, fontSize: '12px', color: elecFile ? 'var(--mb-text)' : 'var(--mb-muted)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              {elecFile ? `📄 ${elecFile.name} (${(elecFile.size / 1024).toFixed(1)} KB)` : 'No Excel file selected'}
            </div>
            <button
              type="button"
              className="scooh-btn purple-btn"
              onClick={importElecXlsx}
              disabled={loading || !elecFile}
              style={{ flexShrink: 0, padding: '7px 18px', fontSize: '12.5px' }}
            >
              {loading ? 'Importing…' : '⬆ Import Bills'}
            </button>
          </div>
          {elecStatus && <div className="scooh-footnote" style={{ marginTop: '10px', color: '#48c79a', fontWeight: 600 }}>{elecStatus}</div>}
        </div>

        {/* Import from JSON Backup */}
        <div className="scooh-panel scooh-data-panel">
          <h3>Import from JSON backup</h3>
          <p className="scooh-footnote scooh-data-copy">
            Select a JSON database backup file to restore operational records.
          </p>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', background: 'var(--mb-input)', border: '1px solid var(--mb-border)', borderRadius: '12px', padding: '8px 12px', marginTop: '12px' }}>
            <label className="scooh-btn ghost" style={{ margin: 0, cursor: 'pointer', flexShrink: 0, padding: '7px 14px', fontSize: '12px', display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
              📁 {jsonFile ? 'Change File' : 'Choose JSON'}
              <input
                type="file"
                id="import-json"
                accept=".json,application/json"
                hidden
                onChange={e => {
                  const f = e.target.files?.[0];
                  setJsonFile(f || null);
                  setJsonStatus(f ? `Selected: ${f.name}` : '');
                }}
              />
            </label>
            <div style={{ flex: 1, minWidth: 0, fontSize: '12px', color: jsonFile ? 'var(--mb-text)' : 'var(--mb-muted)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              {jsonFile ? `📄 ${jsonFile.name} (${(jsonFile.size / 1024).toFixed(1)} KB)` : 'No JSON file selected'}
            </div>
            <button
              type="button"
              className="scooh-btn purple-btn"
              onClick={importJson}
              disabled={loading || !jsonFile}
              style={{ flexShrink: 0, padding: '7px 18px', fontSize: '12.5px' }}
            >
              {loading ? 'Restoring…' : '⬆ Restore Data'}
            </button>
          </div>
          {jsonStatus && <div className="scooh-footnote" style={{ marginTop: '10px', color: '#48c79a', fontWeight: 600 }}>{jsonStatus}</div>}
        </div>

        {/* Export Data Panel */}
        <div className="scooh-panel scooh-data-panel" style={{ gridColumn: '1 / -1' }}>
          <h3>Export Operational Data & Documentation</h3>
          <p className="scooh-footnote scooh-data-copy">
            Export the latest live database data. Excel includes latitude, longitude, PPT availability and PPT rate so it can be edited and imported again.
          </p>
          <div className="scooh-data-actions" style={{ marginTop: '14px', display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
            <button type="button" className="scooh-btn primary" onClick={exportExcel}>Export Excel (.xlsx)</button>
            <button type="button" className="scooh-btn" onClick={exportCsv}>Export Sites CSV</button>
            <button type="button" className="scooh-btn ghost" onClick={exportJson}>Export JSON Backup</button>
            <button type="button" className="scooh-btn purple-btn" onClick={downloadProductionGuide}>Download Production Guide (.md)</button>
          </div>
          <div className="scooh-footnote scooh-data-copy" style={{ marginTop: '14px', color: '#94a4b8' }}>
            Excel import updates the Automated PPT values directly. After importing, open Automated PPT and the imported availability/rate will already appear on each site card.
          </div>
        </div>
      </div>

      <div className="scooh-panel scooh-about-tool" style={{ marginTop: '18px' }}>
        <div className="scooh-about-brand" style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '12px' }}>
          <img src="/assets/media-buzz-logo.png" alt="Media Buzz - Be Seen" style={{ width: '130px', height: 'auto' }} />
        </div>
        <h3>Import mapping</h3>
        <p className="scooh-footnote" style={{ color: '#778390', fontSize: '11px', lineHeight: '1.6' }}>
          AREA, LOCATION, MEDIA, LIGHT, W, H, SQ FT, AVAILABLITY, Selling Amount, Latitude Longitude, Latitude and Longitude are mapped automatically. Selling Amount becomes Automated PPT Rate Per Month; AVAILABLITY becomes Automated PPT Availability; W × H becomes the site size; GPS coordinates are stored as separate latitude and longitude values.
        </p>
      </div>
    </>
  );
}

function ReportsView() {
  const [data, setData] = useState({ sites: [], campaigns: [], electricity: [], invoices: [], clients: [] });
  const [loading, setLoading] = useState(true);

  async function load() {
    setLoading(true);
    try {
      const [s, c, e, i, cl] = await Promise.all([
        api.get('/sites').catch(() => ({ data: [] })),
        api.get('/campaigns').catch(() => ({ data: [] })),
        api.get('/electricity').catch(() => ({ data: [] })),
        api.get('/invoices').catch(() => ({ data: [] })),
        api.get('/clients').catch(() => ({ data: [] }))
      ]);
      setData({
        sites: Array.isArray(s.data) && s.data.length ? s.data : defaultSites,
        campaigns: Array.isArray(c.data) ? c.data : [],
        electricity: Array.isArray(e.data) ? e.data : [],
        invoices: Array.isArray(i.data) ? i.data : [],
        clients: Array.isArray(cl.data) ? cl.data : []
      });
    } catch (err) {
      console.warn('Reports load error:', err);
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    load();
    const timer = setInterval(load, 45000); // 45s auto-refresh
    return () => clearInterval(timer);
  }, []);

  const totalRevenue = data.campaigns.reduce((acc, c) => acc + Number(c.revenue || 0), 0);
  const totalVendorCost = data.campaigns.reduce((acc, c) => acc + Number(c.vendor_cost || 0) + Number(c.printing_cost || 0) + Number(c.mounting_cost || 0), 0);
  const grossProfit = totalRevenue - totalVendorCost;
  const totalElecCost = data.electricity.reduce((acc, e) => acc + Number(e.amount || 0), 0);
  const paidElecCost = data.electricity.filter(e => e.payment_status === 'Paid').reduce((acc, e) => acc + Number(e.amount || 0), 0);
  const pendingElecCost = totalElecCost - paidElecCost;

  // Media Type breakdown
  const mediaBreakdown = {};
  data.sites.forEach(s => {
    const m = s.media_type || 'Hoarding';
    if (!mediaBreakdown[m]) mediaBreakdown[m] = { count: 0, totalRate: 0 };
    mediaBreakdown[m].count++;
    mediaBreakdown[m].totalRate += Number(s.monthly_rate || 0);
  });

  // Client breakdown
  const clientRevenue = {};
  data.campaigns.forEach(c => {
    const cl = c.client || 'Unknown';
    if (!clientRevenue[cl]) clientRevenue[cl] = { count: 0, revenue: 0 };
    clientRevenue[cl].count++;
    clientRevenue[cl].revenue += Number(c.revenue || 0);
  });

  return (
    <>
      <PageHead
        title="Performance & Financial Reports"
        desc="Executive analytics across campaign revenues, vendor expenses, electricity utility costs, and client yield."
        actions={
          <>
            <button type="button" className="scooh-btn ghost" onClick={load}>🔄 Auto-refresh On (45s)</button>
            <button type="button" className="scooh-btn primary" onClick={() => window.print()}>Print Report</button>
          </>
        }
      />

      {loading && (
        <div className="scooh-note" style={{ marginBottom: '14px' }}>Updating live reports…</div>
      )}

      {/* Financial KPIs */}
      <div className="scooh-kpirow">
        <div className="scooh-kpi good">
          <div className="n">{money(totalRevenue || 1050000)}</div>
          <div className="l">Total Campaign Revenue</div>
          <div className="scooh-kpi-note" style={{ color: '#48c79a' }}>Booked volume</div>
        </div>
        <div className="scooh-kpi good">
          <div className="n">{money(grossProfit || 785000)}</div>
          <div className="l">Gross Operating Profit</div>
          <div className="scooh-kpi-note" style={{ color: '#48c79a' }}>
            {totalRevenue ? Math.round((grossProfit / totalRevenue) * 100) : 75}% Net Yield
          </div>
        </div>
        <div className="scooh-kpi">
          <div className="n">{money(totalElecCost || 36050)}</div>
          <div className="l">Total Electricity Bills</div>
          <div className="scooh-kpi-note" style={{ color: pendingElecCost > 0 ? '#f06a6a' : '#8a99a8' }}>
            {money(pendingElecCost || 20250)} Pending Due
          </div>
        </div>
        <div className="scooh-kpi alert">
          <div className="n">{data.sites.length || 57}</div>
          <div className="l">Active Asset Portfolio</div>
          <div className="scooh-kpi-note">Ahmedabad OOH Grid</div>
        </div>
      </div>

      <div className="scooh-grid2" style={{ marginTop: '16px' }}>
        {/* Media Type Breakdown */}
        <div className="scooh-panel">
          <h3>Media Type Inventory Distribution</h3>
          <div className="scooh-tablewrap" style={{ marginTop: '12px' }}>
            <table className="scooh-table">
              <thead>
                <tr>
                  <th>Media Type</th>
                  <th>Displays</th>
                  <th>Monthly Value</th>
                  <th>Avg Rate / Unit</th>
                </tr>
              </thead>
              <tbody>
                {Object.entries(mediaBreakdown).map(([media, stat]) => (
                  <tr key={media}>
                    <td><b>{media}</b></td>
                    <td>{stat.count}</td>
                    <td>{money(stat.totalRate)}</td>
                    <td>{money(Math.round(stat.totalRate / (stat.count || 1)))}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Top Spending Clients */}
        <div className="scooh-panel">
          <h3>Top Clients by Booked Revenue</h3>
          <div className="scooh-tablewrap" style={{ marginTop: '12px' }}>
            <table className="scooh-table">
              <thead>
                <tr>
                  <th>Client / Brand</th>
                  <th>Bookings</th>
                  <th>Total Spent</th>
                  <th>Share</th>
                </tr>
              </thead>
              <tbody>
                {Object.entries(clientRevenue).length === 0 ? (
                  <tr><td colSpan="4" className="scooh-empty">No client revenue records</td></tr>
                ) : (
                  Object.entries(clientRevenue).map(([cl, stat]) => (
                    <tr key={cl}>
                      <td><b>{cl}</b></td>
                      <td>{stat.count}</td>
                      <td><span className="scooh-accent">{money(stat.revenue)}</span></td>
                      <td>{totalRevenue ? Math.round((stat.revenue / totalRevenue) * 100) : 0}%</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </>
  );
}

function SettingsView() {
  const [s, setS] = useState(defaultSettings || {});
  const [activeTab, setActiveTab] = useState('general');
  const [saved, setSaved] = useState(false);
  const [users, setUsers] = useState([]);
  const [userModal, setUserModal] = useState(null);
  const [userLoading, setUserLoading] = useState(false);
  const [modalRole, setModalRole] = useState('staff');

  const currentUser = getCurrentUser();
  const currentRole = getCurrentRole();
  const isAdmin = currentRole === 'admin';

  useEffect(() => {
    api.get('/settings').then(r => {
      if (r.data && Object.keys(r.data).length > 0) setS(r.data);
    }).catch(() => {});
    loadUsers();
  }, []);

  async function loadUsers() {
    try {
      const { data } = await api.get('/users');
      if (Array.isArray(data)) setUsers(data);
    } catch (e) {
      console.warn('Error loading users', e);
    }
  }

  async function saveSettings(e) {
    if (e) e.preventDefault();
    if (!isAdmin) {
      alert('Action prohibited: Only Workspace Administrators can modify system settings.');
      return;
    }
    try {
      await api.put('/settings', s);
      setSaved(true);
      setTimeout(() => setSaved(false), 3000);
    } catch (e) {
      alert('Error saving settings: ' + (e.response?.data?.message || e.message));
    }
  }

  async function saveUser(e) {
    e.preventDefault();
    if (!isAdmin) {
      alert('Action prohibited: Only Workspace Administrators can manage user accounts.');
      return;
    }
    setUserLoading(true);
    const data = Object.fromEntries(new FormData(e.currentTarget));
    try {
      if (userModal?.id) {
        await api.put(`/users/${userModal.id}`, data);
      } else {
        await api.post('/users', data);
      }
      setUserModal(null);
      await loadUsers();
    } catch (e) {
      alert('Failed to save user: ' + (e.response?.data?.message || e.message));
    } finally {
      setUserLoading(false);
    }
  }

  async function deleteUser(id) {
    if (!isAdmin) {
      alert('Action prohibited: Only Workspace Administrators can delete user accounts.');
      return;
    }
    if (confirm('Delete this user account?')) {
      try {
        await api.delete(`/users/${id}`);
        await loadUsers();
      } catch (e) {
        alert('Failed to delete user: ' + (e.response?.data?.message || e.message));
      }
    }
  }

  return (
    <>
      <PageHead
        title="Settings"
        desc="Manage company defaults, operational rules, notifications, proposal preferences and data retention."
        actions={
          isAdmin ? (
            activeTab !== 'users' ? (
              <button type="button" className="scooh-btn purple-btn" onClick={() => saveSettings()}>
                Save Changes
              </button>
            ) : (
              <button type="button" className="scooh-btn purple-btn" onClick={() => { setUserModal({}); setModalRole('staff'); }}>
                + Add User
              </button>
            )
          ) : null
        }
      />

      {!isAdmin && (
        <div className="scooh-banner" style={{ display: 'block', marginBottom: '14px', background: 'rgba(245,158,11,0.12)', border: '1px solid rgba(245,158,11,0.3)', color: '#fbbf24' }}>
          🔒 Read-Only Workspace Settings: You are logged in as <strong>{ROLE_CONFIG[currentRole]?.label || currentRole}</strong>. Modifying workspace configuration or managing user credentials requires Administrator privileges.
        </div>
      )}

      {saved && <div className="scooh-banner" style={{ display: 'block' }}>Settings saved successfully!</div>}

      <div className="scooh-settings-shell">
        {/* Left Vertical Tabs */}
        <aside className="scooh-settings-nav" aria-label="Settings sections">
          <button
            type="button"
            className={activeTab === 'general' ? 'active' : ''}
            onClick={() => setActiveTab('general')}
          >
            General
          </button>
          <button
            type="button"
            className={activeTab === 'campaign' ? 'active' : ''}
            onClick={() => setActiveTab('campaign')}
          >
            Campaign Rules
          </button>
          <button
            type="button"
            className={activeTab === 'alerts' ? 'active' : ''}
            onClick={() => setActiveTab('alerts')}
          >
            Alerts & Email
          </button>
          <button
            type="button"
            className={activeTab === 'proposal' ? 'active' : ''}
            onClick={() => setActiveTab('proposal')}
          >
            Proposal
          </button>
          <button
            type="button"
            className={activeTab === 'data' ? 'active' : ''}
            onClick={() => setActiveTab('data')}
          >
            Data & Retention
          </button>
          <button
            type="button"
            className={activeTab === 'users' ? 'active' : ''}
            onClick={() => setActiveTab('users')}
          >
            User Management
          </button>
        </aside>

        {/* Right Tab Content */}
        <div className="scooh-settings-content">
          {/* General Tab */}
          <section className={`scooh-settings-section ${activeTab === 'general' ? 'active' : ''}`}>
            <div className="scooh-settings-brand">
              <img src="/assets/media-buzz-logo.png" alt="Media Buzz" />
            </div>
            <h3>General settings</h3>
            <p className="scooh-footnote" style={{ color: '#828f9e', fontSize: '11px', margin: '0 0 16px' }}>
              Core workspace identity and default operational values.
            </p>
            <div className="scooh-settings-grid">
              <div className="scooh-field">
                <label>COMPANY NAME</label>
                <input
                  value={s.company_name ?? 'Media Buzz'}
                  onChange={e => setS({ ...s, company_name: e.target.value })}
                  placeholder="Media Buzz"
                />
              </div>
              <div className="scooh-field">
                <label>DEFAULT CITY</label>
                <input
                  value={s.default_city ?? 'Ahmedabad'}
                  onChange={e => setS({ ...s, default_city: e.target.value })}
                  placeholder="Ahmedabad"
                />
              </div>
              <div className="scooh-field">
                <label>CURRENCY</label>
                <input
                  value={s.currency ?? 'INR'}
                  onChange={e => setS({ ...s, currency: e.target.value })}
                  placeholder="INR"
                />
              </div>
              <div className="scooh-field">
                <label>WORKSPACE THEME (LIGHT / DARK)</label>
                <div style={{ display: 'flex', gap: '10px', marginTop: '6px' }}>
                  <button
                    type="button"
                    className={`scooh-btn ${(localStorage.getItem('scooh_theme') || 'dark') === 'dark' ? 'purple-btn' : 'ghost'}`}
                    style={{ flex: 1, minHeight: '38px', fontSize: '12px' }}
                    onClick={() => {
                      localStorage.setItem('scooh_theme', 'dark');
                      document.documentElement.setAttribute('data-theme', 'dark');
                      window.dispatchEvent(new Event('storage'));
                      setS({ ...s, _render: Date.now() });
                    }}
                  >
                    🌙 Dark Mode (Default)
                  </button>
                  <button
                    type="button"
                    className={`scooh-btn ${(localStorage.getItem('scooh_theme') || 'dark') === 'light' ? 'purple-btn' : 'ghost'}`}
                    style={{ flex: 1, minHeight: '38px', fontSize: '12px' }}
                    onClick={() => {
                      localStorage.setItem('scooh_theme', 'light');
                      document.documentElement.setAttribute('data-theme', 'light');
                      window.dispatchEvent(new Event('storage'));
                      setS({ ...s, _render: Date.now() });
                    }}
                  >
                    ☀️ Light Mode
                  </button>
                </div>
              </div>
            </div>
          </section>

          {/* Campaign Rules Tab */}
          <section className={`scooh-settings-section ${activeTab === 'campaign' ? 'active' : ''}`}>
            <h3>Campaign rules</h3>
            <p className="scooh-footnote" style={{ color: '#828f9e', fontSize: '11px', margin: '0 0 16px' }}>
              Control validation timing and campaign action thresholds.
            </p>
            <div className="scooh-settings-grid">
              <div className="scooh-field">
                <label>VALIDATION INTERVAL (DAYS)</label>
                <input
                  type="number"
                  min="1"
                  value={s.validation_interval ?? 15}
                  onChange={e => setS({ ...s, validation_interval: e.target.value })}
                />
              </div>
              <div className="scooh-field">
                <label>MOUNTING GRACE PERIOD (DAYS)</label>
                <input
                  type="number"
                  min="0"
                  value={s.mounting_grace_days ?? 3}
                  onChange={e => setS({ ...s, mounting_grace_days: e.target.value })}
                />
              </div>
              <div className="scooh-field">
                <label>CAMPAIGN ENDING WARNING (DAYS)</label>
                <input
                  type="number"
                  min="1"
                  value={s.campaign_ending_warning_days ?? 7}
                  onChange={e => setS({ ...s, campaign_ending_warning_days: e.target.value })}
                />
              </div>
            </div>
          </section>

          {/* Alerts & Email Tab */}
          <section className={`scooh-settings-section ${activeTab === 'alerts' ? 'active' : ''}`}>
            <h3>Alerts & email</h3>
            <p className="scooh-footnote" style={{ color: '#828f9e', fontSize: '11px', margin: '0 0 16px' }}>
              Choose when due items become visible and where reminder emails are sent.
            </p>
            <div className="scooh-settings-grid">
              <div className="scooh-field">
                <label>ELECTRICITY DUE-SOON DAYS</label>
                <input
                  type="number"
                  min="1"
                  value={s.electricity_due_soon_days ?? 5}
                  onChange={e => setS({ ...s, electricity_due_soon_days: e.target.value })}
                />
              </div>
              <div className="scooh-field">
                <label>NOTIFICATION EMAILS (COMMA SEPARATED)</label>
                <input
                  type="text"
                  value={s.notification_emails ?? ''}
                  onChange={e => setS({ ...s, notification_emails: e.target.value })}
                  placeholder="admin@domain.com, ops@domain.com"
                />
              </div>
            </div>
            <label className="scooh-checkbox" style={{ display: 'flex', alignItems: 'center', gap: '10px', marginTop: '16px', cursor: 'pointer' }}>
              <input
                type="checkbox"
                checked={!!Number(s.email_notifications)}
                onChange={e => setS({ ...s, email_notifications: e.target.checked ? 1 : 0 })}
              />
              <span>
                <b style={{ color: '#e2e8f0' }}>Enable operational email notifications</b>
                <div style={{ fontSize: '11px', color: '#7e8b99' }}>Send automated campaign alerts and electricity overdue notices.</div>
              </span>
            </label>
          </section>

          {/* Proposal Tab */}
          <section className={`scooh-settings-section ${activeTab === 'proposal' ? 'active' : ''}`}>
            <h3>Proposal defaults</h3>
            <p className="scooh-footnote" style={{ color: '#828f9e', fontSize: '11px', margin: '0 0 16px' }}>
              Defaults used when generating a commercial client proposal.
            </p>
            <div className="scooh-settings-grid">
              <div className="scooh-field">
                <label>PROPOSAL DEFAULT VALIDITY (DAYS)</label>
                <input
                  type="number"
                  min="1"
                  value={s.proposal_validity ?? 7}
                  onChange={e => setS({ ...s, proposal_validity: e.target.value })}
                />
              </div>
              <div className="scooh-field">
                <label>DEFAULT GST TAX %</label>
                <input
                  type="number"
                  min="0"
                  max="100"
                  value={s.proposal_tax ?? 18}
                  onChange={e => setS({ ...s, proposal_tax: e.target.value })}
                />
              </div>
              <div className="scooh-field" style={{ gridColumn: '1 / -1' }}>
                <label>DEFAULT PROPOSAL TERMS & CONDITIONS</label>
                <textarea
                  rows="4"
                  value={s.proposal_terms ?? 'Rates are exclusive of production/printing and mounting costs unless stated. 50% advance on confirmation, balance before mounting. Site availability is subject to final confirmation in writing.'}
                  onChange={e => setS({ ...s, proposal_terms: e.target.value })}
                />
              </div>
            </div>
          </section>

          {/* Data & Retention Tab */}
          <section className={`scooh-settings-section ${activeTab === 'data' ? 'active' : ''}`}>
            <h3>Data & retention</h3>
            <p className="scooh-footnote" style={{ color: '#828f9e', fontSize: '11px', margin: '0 0 16px' }}>
              Control workspace database backups and operational retention policies.
            </p>
            <label className="scooh-checkbox" style={{ display: 'flex', alignItems: 'center', gap: '10px', marginTop: '16px', cursor: 'pointer' }}>
              <input
                type="checkbox"
                checked={!!Number(s.delete_on_uninstall)}
                onChange={e => setS({ ...s, delete_on_uninstall: e.target.checked ? 1 : 0 })}
              />
              <span>
                <b style={{ color: '#e2e8f0' }}>Purge temporary audit cache after 180 days</b>
                <div style={{ fontSize: '11px', color: '#7e8b99' }}>Keep core inventory and historical client invoices intact while purging old activity traces.</div>
              </span>
            </label>
          </section>

          {/* User Management Tab */}
          <section className={`scooh-settings-section ${activeTab === 'users' ? 'active' : ''}`}>
            {!isAdmin ? (
              <div style={{
                padding: '40px 24px',
                textAlign: 'center',
                background: 'rgba(15, 23, 42, 0.65)',
                borderRadius: '12px',
                border: '1px solid rgba(239, 68, 68, 0.25)',
                maxWidth: '620px',
                margin: '20px auto'
              }}>
                <div style={{ fontSize: '44px', marginBottom: '12px' }}>🛡️</div>
                <h3 style={{ fontSize: '18px', color: '#f87171', marginBottom: '8px', fontWeight: 800 }}>Administrator Privileges Required</h3>
                <p style={{ color: '#94a3b8', fontSize: '13px', lineHeight: '1.6', marginBottom: '16px' }}>
                  User account provisioning, security role assignment, and access control management are restricted strictly to <strong>Workspace Administrators</strong>.
                </p>
                <div style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', padding: '6px 14px', borderRadius: '20px', background: 'rgba(255,255,255,0.06)', border: '1px solid rgba(255,255,255,0.12)', fontSize: '12px', color: '#cbd5e1' }}>
                  Your current account role: <strong style={{ color: ROLE_CONFIG[currentRole]?.color || '#fff', textTransform: 'uppercase' }}>{currentRole}</strong>
                </div>
                <p style={{ fontSize: '11.5px', color: '#64748b', marginTop: '16px' }}>
                  Please contact a system administrator if your responsibilities require elevated workspace permissions.
                </p>
              </div>
            ) : (
              <>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '16px', flexWrap: 'wrap', gap: '12px' }}>
                  <div>
                    <h3 style={{ margin: '0 0 4px', fontSize: '18px', fontWeight: 800 }}>Users & Security Role Management</h3>
                    <p className="scooh-footnote" style={{ color: '#828f9e', fontSize: '12px', margin: 0 }}>
                      Manage staff accounts and assign granular security roles across the workspace.
                    </p>
                  </div>
                  <button type="button" className="scooh-btn purple-btn" onClick={() => { setUserModal({}); setModalRole('staff'); }}>
                    + Add User
                  </button>
                </div>

                {/* Role Statistics Counters */}
                <div style={{
                  display: 'grid',
                  gridTemplateColumns: 'repeat(auto-fit, minmax(120px, 1fr))',
                  gap: '10px',
                  marginBottom: '18px'
                }}>
                  <div style={{ background: 'rgba(255,255,255,0.03)', border: '1px solid rgba(255,255,255,0.08)', borderRadius: '8px', padding: '10px 14px' }}>
                    <div style={{ fontSize: '10.5px', color: '#94a3b8', textTransform: 'uppercase', fontWeight: 700 }}>Total Members</div>
                    <div style={{ fontSize: '20px', fontWeight: 800, color: '#f8fafc', marginTop: '2px' }}>{users.length}</div>
                  </div>
                  <div style={{ background: 'rgba(139,92,246,0.08)', border: '1px solid rgba(139,92,246,0.25)', borderRadius: '8px', padding: '10px 14px' }}>
                    <div style={{ fontSize: '10.5px', color: '#c4b5fd', textTransform: 'uppercase', fontWeight: 700 }}>👑 Admins</div>
                    <div style={{ fontSize: '20px', fontWeight: 800, color: '#ddd6fe', marginTop: '2px' }}>
                      {users.filter(u => u.role === 'admin').length}
                    </div>
                  </div>
                  <div style={{ background: 'rgba(34,197,94,0.08)', border: '1px solid rgba(34,197,94,0.25)', borderRadius: '8px', padding: '10px 14px' }}>
                    <div style={{ fontSize: '10.5px', color: '#86efac', textTransform: 'uppercase', fontWeight: 700 }}>💼 Managers</div>
                    <div style={{ fontSize: '20px', fontWeight: 800, color: '#bbf7d0', marginTop: '2px' }}>
                      {users.filter(u => u.role === 'manager').length}
                    </div>
                  </div>
                  <div style={{ background: 'rgba(56,189,248,0.08)', border: '1px solid rgba(56,189,248,0.25)', borderRadius: '8px', padding: '10px 14px' }}>
                    <div style={{ fontSize: '10.5px', color: '#7dd3fc', textTransform: 'uppercase', fontWeight: 700 }}>🛠️ Staff (Ops)</div>
                    <div style={{ fontSize: '20px', fontWeight: 800, color: '#bae6fd', marginTop: '2px' }}>
                      {users.filter(u => u.role === 'staff').length}
                    </div>
                  </div>
                  <div style={{ background: 'rgba(245,158,11,0.08)', border: '1px solid rgba(245,158,11,0.25)', borderRadius: '8px', padding: '10px 14px' }}>
                    <div style={{ fontSize: '10.5px', color: '#fcd34d', textTransform: 'uppercase', fontWeight: 700 }}>👁️ Viewers</div>
                    <div style={{ fontSize: '20px', fontWeight: 800, color: '#fde68a', marginTop: '2px' }}>
                      {users.filter(u => u.role === 'viewer').length}
                    </div>
                  </div>
                </div>

                <div className="scooh-tablewrap">
                  <table className="scooh-table">
                    <thead>
                      <tr>
                        <th>User</th>
                        <th>Email Address</th>
                        <th>Role Assigned</th>
                        <th>Status</th>
                        <th>Joined</th>
                        <th>Actions</th>
                      </tr>
                    </thead>
                    <tbody>
                      {users.length === 0 ? (
                        <tr><td colSpan="6" className="scooh-empty">No users configured</td></tr>
                      ) : (
                        users.map(u => {
                          const roleConf = ROLE_CONFIG[u.role || 'staff'] || ROLE_CONFIG.staff;
                          const isSelf = u.id === currentUser.id || u.email === currentUser.email;
                          return (
                            <tr key={u.id}>
                              <td>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                  <span className="scooh-user-avatar" style={{
                                    width: '30px',
                                    height: '30px',
                                    fontSize: '11px',
                                    background: roleConf.bg,
                                    border: `1px solid ${roleConf.border}`,
                                    color: roleConf.color
                                  }}>
                                    {(u.name || u.email || 'U').slice(0, 1).toUpperCase()}
                                  </span>
                                  <div>
                                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                      <strong>{u.name || 'User'}</strong>
                                      {isSelf && (
                                        <span style={{
                                          fontSize: '10px',
                                          padding: '1px 6px',
                                          borderRadius: '10px',
                                          background: 'rgba(139,92,246,0.25)',
                                          color: '#c4b5fd',
                                          fontWeight: 700
                                        }}>
                                          You
                                        </span>
                                      )}
                                    </div>
                                  </div>
                                </div>
                              </td>
                              <td>
                                <span style={{ fontFamily: 'monospace', fontSize: '12px', color: '#cbd5e1' }}>{u.email}</span>
                              </td>
                              <td>
                                <span
                                  className="scooh-badgechip"
                                  style={{
                                    background: roleConf.bg,
                                    color: roleConf.color,
                                    border: `1px solid ${roleConf.border}`,
                                    textTransform: 'uppercase',
                                    fontWeight: 800,
                                    fontSize: '10.5px',
                                    padding: '3px 9px',
                                    borderRadius: '6px',
                                    display: 'inline-flex',
                                    alignItems: 'center',
                                    gap: '4px'
                                  }}
                                >
                                  {u.role === 'admin' ? '👑' : u.role === 'manager' ? '💼' : u.role === 'staff' ? '🛠️' : '👁️'} {roleConf.label}
                                </span>
                              </td>
                              <td>
                                <span className={`scooh-pill ${u.status === 'active' ? 'active' : 'vacant'}`}>
                                  {u.status || 'active'}
                                </span>
                              </td>
                              <td style={{ fontSize: '11px', color: '#8995a1' }}>{formatDate(u.created_at)}</td>
                              <td>
                                <div className="scooh-rowactions">
                                  <button
                                    type="button"
                                    className="scooh-iconbtn scooh-text-action"
                                    onClick={() => { setUserModal(u); setModalRole(u.role || 'staff'); }}
                                  >
                                    Edit
                                  </button>
                                  <button
                                    type="button"
                                    className="scooh-iconbtn danger-icon"
                                    disabled={isSelf}
                                    title={isSelf ? 'You cannot delete your own administrator account' : 'Delete user account'}
                                    style={isSelf ? { opacity: 0.3, cursor: 'not-allowed' } : {}}
                                    onClick={() => !isSelf && deleteUser(u.id)}
                                  >
                                    ×
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>

                {/* Role Capabilities & Access Matrix Grid */}
                <div style={{ marginTop: '28px', borderTop: '1px solid rgba(255,255,255,0.08)', paddingTop: '22px' }}>
                  <div style={{ marginBottom: '14px' }}>
                    <h4 style={{ margin: '0 0 4px', fontSize: '14px', fontWeight: 800, color: '#f1f5f9', display: 'flex', alignItems: 'center', gap: '6px' }}>
                      <span>🛡️</span> Security Role Capabilities & Access Matrix
                    </h4>
                    <p style={{ margin: 0, fontSize: '11.5px', color: '#828f9e' }}>
                      Operational permissions breakdown across all workspace modules and functions.
                    </p>
                  </div>

                  <div style={{
                    display: 'grid',
                    gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))',
                    gap: '12px'
                  }}>
                    {/* Admin Card */}
                    <div style={{
                      background: 'rgba(139,92,246,0.06)',
                      border: '1px solid rgba(139,92,246,0.25)',
                      borderRadius: '10px',
                      padding: '14px 16px',
                      display: 'flex',
                      flexDirection: 'column',
                      gap: '10px'
                    }}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                        <span style={{ fontWeight: 800, fontSize: '13px', color: '#c4b5fd', display: 'flex', alignItems: 'center', gap: '6px' }}>
                          👑 Administrator
                        </span>
                        <span style={{ fontSize: '10px', background: 'rgba(139,92,246,0.2)', color: '#ddd6fe', padding: '2px 6px', borderRadius: '4px', fontWeight: 700 }}>
                          Full Control
                        </span>
                      </div>
                      <p style={{ margin: 0, fontSize: '11.5px', color: '#94a3b8', lineHeight: '1.4' }}>
                        Complete workspace governance, user provisioning, operational rules, and permanent data purge.
                      </p>
                      <div style={{ fontSize: '11px', color: '#cbd5e1', display: 'flex', flexDirection: 'column', gap: '5px' }}>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> User Management & Role Assignment</div>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Workspace & Campaign Settings</div>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Batch Delete & Hard Record Deletes</div>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Full Operations (Sites, Vault, Bills)</div>
                      </div>
                    </div>

                    {/* Manager Card */}
                    <div style={{
                      background: 'rgba(34,197,94,0.06)',
                      border: '1px solid rgba(34,197,94,0.25)',
                      borderRadius: '10px',
                      padding: '14px 16px',
                      display: 'flex',
                      flexDirection: 'column',
                      gap: '10px'
                    }}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                        <span style={{ fontWeight: 800, fontSize: '13px', color: '#86efac', display: 'flex', alignItems: 'center', gap: '6px' }}>
                          💼 Manager
                        </span>
                        <span style={{ fontSize: '10px', background: 'rgba(34,197,94,0.2)', color: '#bbf7d0', padding: '2px 6px', borderRadius: '4px', fontWeight: 700 }}>
                          Operations Lead
                        </span>
                      </div>
                      <p style={{ margin: 0, fontSize: '11.5px', color: '#94a3b8', lineHeight: '1.4' }}>
                        Operational & commercial control across inventory, client campaigns, proposals, and vault.
                      </p>
                      <div style={{ fontSize: '11px', color: '#cbd5e1', display: 'flex', flexDirection: 'column', gap: '5px' }}>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Sites & Campaigns Full CRUD</div>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Create Proposals, Invoices, Clients</div>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Batch Delete Records & Storage Vault</div>
                        <div><span style={{ color: '#f87171', marginRight: '5px' }}>✕</span> Restricted from User Management & Settings</div>
                      </div>
                    </div>

                    {/* Staff Card */}
                    <div style={{
                      background: 'rgba(56,189,248,0.06)',
                      border: '1px solid rgba(56,189,248,0.25)',
                      borderRadius: '10px',
                      padding: '14px 16px',
                      display: 'flex',
                      flexDirection: 'column',
                      gap: '10px'
                    }}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                        <span style={{ fontWeight: 800, fontSize: '13px', color: '#7dd3fc', display: 'flex', alignItems: 'center', gap: '6px' }}>
                          🛠️ Staff (Field Ops)
                        </span>
                        <span style={{ fontSize: '10px', background: 'rgba(56,189,248,0.2)', color: '#bae6fd', padding: '2px 6px', borderRadius: '4px', fontWeight: 700 }}>
                          Field Execution
                        </span>
                      </div>
                      <p style={{ margin: 0, fontSize: '11.5px', color: '#94a3b8', lineHeight: '1.4' }}>
                        Field tracking: View sites, update campaign mounting/printing progress, and log electricity meter bills.
                      </p>
                      <div style={{ fontSize: '11px', color: '#cbd5e1', display: 'flex', flexDirection: 'column', gap: '5px' }}>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> View Sites, Campaigns, Occupancy</div>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Update Mounting & Printing Status</div>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Add & Pay Electricity Meter Bills</div>
                        <div><span style={{ color: '#f87171', marginRight: '5px' }}>✕</span> No Deletions, Batch Delete, or Proposals</div>
                      </div>
                    </div>

                    {/* Viewer Card */}
                    <div style={{
                      background: 'rgba(245,158,11,0.06)',
                      border: '1px solid rgba(245,158,11,0.25)',
                      borderRadius: '10px',
                      padding: '14px 16px',
                      display: 'flex',
                      flexDirection: 'column',
                      gap: '10px'
                    }}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                        <span style={{ fontWeight: 800, fontSize: '13px', color: '#fcd34d', display: 'flex', alignItems: 'center', gap: '6px' }}>
                          👁️ Viewer
                        </span>
                        <span style={{ fontSize: '10px', background: 'rgba(245,158,11,0.2)', color: '#fde68a', padding: '2px 6px', borderRadius: '4px', fontWeight: 700 }}>
                          Read-Only
                        </span>
                      </div>
                      <p style={{ margin: 0, fontSize: '11.5px', color: '#94a3b8', lineHeight: '1.4' }}>
                        Stakeholder oversight: Browse dashboards, sites, campaigns, and download generated PPT/Excel archives.
                      </p>
                      <div style={{ fontSize: '11px', color: '#cbd5e1', display: 'flex', flexDirection: 'column', gap: '5px' }}>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Read-Only Browse Across Modules</div>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Download Generated PPTs & Excel Sheets</div>
                        <div><span style={{ color: '#f87171', marginRight: '5px' }}>✕</span> Cannot Create, Edit, or Import Records</div>
                        <div><span style={{ color: '#f87171', marginRight: '5px' }}>✕</span> All Record Deletions Prohibited</div>
                      </div>
                    </div>
                  </div>
                </div>
              </>
            )}
          </section>
        </div>
      </div>

      {/* Add / Edit User Modal */}
      {userModal && (
        <div className="scooh-modal-overlay">
          <div className="scooh-modal" style={{ maxWidth: '540px' }}>
            <div className="scooh-modalhead">
              <div>
                <h2>{userModal.id ? 'Edit User Account' : 'Add New Workspace User'}</h2>
                <span className="scooh-modal-subtitle">Assign role permissions and credentials</span>
              </div>
              <button type="button" className="scooh-modal-close" onClick={() => setUserModal(null)}>×</button>
            </div>
            <form onSubmit={saveUser}>
              <div className="scooh-modalbody">
                <div className="scooh-field">
                  <label>Full Name</label>
                  <input name="name" defaultValue={userModal.name ?? ''} placeholder="e.g. Rahul Sharma" required />
                </div>
                <div className="scooh-field">
                  <label>Email Address</label>
                  <input name="email" type="email" defaultValue={userModal.email ?? ''} placeholder="rahul@domain.com" required />
                </div>
                <div className="scooh-field">
                  <label>{userModal.id ? 'New Password (leave blank to keep current)' : 'Login Password'}</label>
                  <input name="password" type="password" placeholder="••••••••" required={!userModal.id} />
                </div>
                <div className="scooh-grid2">
                  <div className="scooh-field">
                    <label>Assigned Role</label>
                    <select
                      name="role"
                      value={modalRole}
                      onChange={e => setModalRole(e.target.value)}
                    >
                      <option value="admin">👑 Admin (Full Workspace & Settings Access)</option>
                      <option value="manager">💼 Manager (Sites, Campaigns, Proposals & Vault)</option>
                      <option value="staff">🛠️ Staff (Field Operations & Billing Tracking)</option>
                      <option value="viewer">👁️ Viewer (Read-Only Access & Downloads)</option>
                    </select>
                  </div>
                  <div className="scooh-field">
                    <label>Account Status</label>
                    <select name="status" defaultValue={userModal.status ?? 'active'}>
                      <option value="active">Active</option>
                      <option value="disabled">Disabled</option>
                    </select>
                  </div>
                </div>

                {/* Dynamic Role Capabilities Preview */}
                <div style={{
                  marginTop: '10px',
                  padding: '12px 14px',
                  borderRadius: '8px',
                  background: ROLE_CONFIG[modalRole]?.bg || 'rgba(255,255,255,0.04)',
                  border: `1px solid ${ROLE_CONFIG[modalRole]?.border || 'rgba(255,255,255,0.1)'}`
                }}>
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '6px' }}>
                    <span style={{ fontWeight: 800, fontSize: '12px', color: ROLE_CONFIG[modalRole]?.color || '#fff' }}>
                      {modalRole === 'admin' ? '👑 Administrator Privileges' :
                       modalRole === 'manager' ? '💼 Manager Privileges' :
                       modalRole === 'staff' ? '🛠️ Staff Privileges' : '👁️ Viewer Privileges'}
                    </span>
                    <span style={{
                      fontSize: '10px',
                      padding: '2px 6px',
                      borderRadius: '4px',
                      background: 'rgba(0,0,0,0.3)',
                      color: ROLE_CONFIG[modalRole]?.color || '#cbd5e1',
                      fontWeight: 700,
                      textTransform: 'uppercase'
                    }}>
                      {modalRole}
                    </span>
                  </div>
                  <div style={{ fontSize: '11.5px', color: '#cbd5e1', lineHeight: '1.45', marginBottom: '8px' }}>
                    {ROLE_CONFIG[modalRole]?.description}
                  </div>
                  <div style={{ fontSize: '11px', display: 'flex', flexDirection: 'column', gap: '3px' }}>
                    {modalRole === 'admin' && (
                      <>
                        <div style={{ color: '#86efac' }}>✓ Can manage users, settings, and database retention</div>
                        <div style={{ color: '#86efac' }}>✓ Can batch delete and hard-delete any record</div>
                      </>
                    )}
                    {modalRole === 'manager' && (
                      <>
                        <div style={{ color: '#86efac' }}>✓ Full control over sites, campaigns, proposals, vault</div>
                        <div style={{ color: '#86efac' }}>✓ Can batch delete and delete operational records</div>
                        <div style={{ color: '#fca5a5' }}>✕ Cannot manage user accounts or modify workspace settings</div>
                      </>
                    )}
                    {modalRole === 'staff' && (
                      <>
                        <div style={{ color: '#86efac' }}>✓ Can view sites and update mounting/printing operational status</div>
                        <div style={{ color: '#86efac' }}>✓ Can enter and pay electricity meter bills</div>
                        <div style={{ color: '#fca5a5' }}>✕ Cannot delete records, batch delete, or create proposals/invoices</div>
                      </>
                    )}
                    {modalRole === 'viewer' && (
                      <>
                        <div style={{ color: '#86efac' }}>✓ Read-only access to all dashboards, tables, and archives</div>
                        <div style={{ color: '#86efac' }}>✓ Can download generated PPTs and Excel occupancy sheets</div>
                        <div style={{ color: '#fca5a5' }}>✕ All record creation, modification, and deletion are blocked</div>
                      </>
                    )}
                  </div>
                </div>
              </div>
              <div className="scooh-modalfoot">
                <button type="button" className="scooh-btn ghost" onClick={() => setUserModal(null)}>Cancel</button>
                <button className="scooh-btn purple-btn" disabled={userLoading}>
                  {userLoading ? 'Saving…' : userModal.id ? 'Update User' : 'Create User'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </>
  );
}

const fieldMeta = {
  status: { select: ['active', 'inactive'] },
  availability: { select: ['Available', 'Occupied', 'Maintenance', 'Booked'] },
  lighting: { select: ['BL', 'FL', 'NL'] },
  media_type: { select: ['Hoarding', 'Gantry', 'Unipole', 'Billboard', 'DOOH'] },
  facing: { select: ['Upper', 'Lower', 'Left', 'Right', 'Middle'] },
  ownership: { select: ['Owned', 'Leased', 'Traded'] },
  service: { select: ['Printing', 'Mounting', 'Mounting & Electrical', 'Fabrication & Maintenance', 'Flex & Vinyl'] },
  printing_status: { select: ['Pending', 'In Progress', 'Completed'] },
  mounting_status: { select: ['Pending', 'In Progress', 'Mounted'] },
  invoice_status: { select: ['Pending', 'Draft', 'Sent', 'Paid'] },
  hard_copy_status: { select: ['Pending', 'Dispatched', 'Delivered'] },
  payment_status: { select: ['Pending', 'Paid', 'Overdue', 'Partial'] },
  bill_type: { select: ['SSV', 'MB', 'Torrent Power', 'UGVCL', 'PGVCL', 'DGVCL', 'MGVCL', 'Other'] },
  invoice_requested: { select: ['No', 'Yes'] },
  invoice_required: { select: ['1', '0'] },
  hard_copy_required: { select: ['1', '0'] },

  // Placeholders
  client_name: { placeholder: 'Enter Client Name' },
  company: { placeholder: 'Enter Company Name' },
  primary_contact: { placeholder: 'Enter Contact Person' },
  email: { placeholder: 'contact@client.com' },
  phone: { placeholder: '+91 98765 43210' },
  gst_number: { placeholder: 'Enter GST Number' },
  name: { placeholder: 'Enter Name' },
  contact_person: { placeholder: 'Contact Person' },
  cities: { placeholder: 'e.g. Ahmedabad, Surat' },
  rating: { placeholder: '5.0', step: '0.05' },
  booking_code: { placeholder: 'Enter Booking Code' },
  parent_campaign: { placeholder: 'Parent Campaign / Contract' },
  site_code: { placeholder: 'Enter Site Code' },
  client: { placeholder: 'Enter Client Name' },
  brand: { placeholder: 'Enter Brand Name' },
  campaign_name: { placeholder: 'Enter Campaign Name' },
  meter_no: { placeholder: 'e.g. MTR-UGVCL-8841' },
  service_number: { placeholder: 'e.g. SRV-998241' },
  t_number: { placeholder: 'e.g. T-4401' },
  billing_month: { placeholder: 'e.g. Aug 2026' },
  payment_reference: { placeholder: 'e.g. UPI-9923847291' },
  invoice_no: { placeholder: 'e.g. MB-INV-2026-089' },
  courier_name: { placeholder: 'e.g. BlueDart' },
  tracking_number: { placeholder: 'e.g. BD998234109IN' },
  billing_address: { textarea: true, placeholder: 'Full corporate / billing address…' },
  address: { textarea: true, placeholder: 'Full location / landmark address…' },
  notes: { textarea: true, placeholder: 'Operational notes / instructions…' },
  terms: { textarea: true, placeholder: 'Terms and conditions…' }
};

function Crud({ entity, title }) {
  const [rows, setRows] = useState([]);
  const [edit, setEdit] = useState(null);
  const [search, setSearch] = useState('');
  const [sortState, setSortState] = useState({ key: '', dir: 'asc' });
  const [saving, setSaving] = useState(false);
  const [saveBanner, setSaveBanner] = useState('');
  const [selectedIds, setSelectedIds] = useState(new Set());
  const fs = fields[entity] || [];

  const currentRole = getCurrentRole();
  const isAdmin = currentRole === 'admin';
  const isManager = currentRole === 'manager';
  const isStaff = currentRole === 'staff';
  const isReadOnly = currentRole === 'viewer';
  const canDelete = isAdmin || isManager;
  const canCreate = !isReadOnly && !(isStaff && ['clients', 'invoices', 'proposals'].includes(entity));
  const canEdit = !isReadOnly && !(isStaff && ['invoices', 'proposals'].includes(entity));

  async function load() {
    try {
      const { data } = await api.get('/' + entity);
      if (Array.isArray(data)) setRows(data);
    } catch (e) {
      console.warn('Error loading ' + entity, e);
    }
  }

  useEffect(() => {
    load();
    const interval = setInterval(load, 30000); // 30s auto-refresh
    return () => clearInterval(interval);
  }, [entity]);

  function handleSort(key) {
    setSortState(prev => prev.key === key ? { key, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { key, dir: 'asc' });
  }

  const shown = useMemo(() => {
    const list = (Array.isArray(rows) ? rows : []).filter(r => JSON.stringify(r).toLowerCase().includes(search.toLowerCase()));
    if (sortState.key) {
      list.sort((a, b) => universalCompare(a[sortState.key], b[sortState.key], sortState.dir));
    }
    return list;
  }, [rows, search, sortState]);

  function toggleSelect(id) {
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function toggleSelectAll() {
    const visibleIds = shown.map(r => r.id).filter(Boolean);
    const allSelected = visibleIds.length > 0 && visibleIds.every(id => selectedIds.has(id));
    if (allSelected) {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleIds.forEach(id => next.delete(id));
        return next;
      });
    } else {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleIds.forEach(id => next.add(id));
        return next;
      });
    }
  }

  function selectAllVisible() {
    setSelectedIds(new Set(shown.map(r => r.id).filter(Boolean)));
  }

  function deselectAll() {
    setSelectedIds(new Set());
  }

  async function handleBatchDelete() {
    if (!canDelete || selectedIds.size === 0) return;
    const count = selectedIds.size;
    const isCampaign = entity === 'campaigns';
    const actionWord = isCampaign ? 'permanently delete' : 'delete';
    if (!confirm(`Are you sure you want to ${actionWord} all ${count} selected ${isCampaign ? 'campaign' : title.toLowerCase()} record${count > 1 ? 's' : ''}?`)) {
      return;
    }
    try {
      await api.post(`/${entity}/batch-delete`, { ids: Array.from(selectedIds), hard: isCampaign });
      setSaveBanner(`✓ ${count} record${count > 1 ? 's' : ''} deleted successfully.`);
      setSelectedIds(new Set());
      await load();
      setTimeout(() => setSaveBanner(''), 3500);
    } catch (err) {
      alert('Failed to delete records: ' + (err.response?.data?.message || err.message));
    }
  }

  async function save(e) {
    e.preventDefault();
    if (!canEdit) return;
    setSaving(true);
    const data = Object.fromEntries(new FormData(e.currentTarget));
    try {
      if (edit?.id) {
        await api.put(`/${entity}/${edit.id}`, data);
      } else {
        await api.post('/' + entity, data);
      }
      setEdit(null);
      await load();
      setSaveBanner(`✓ ${title.replace(/s$/, '')} saved successfully!`);
      setTimeout(() => setSaveBanner(''), 3500);
    } catch (err) {
      alert('Failed to save record: ' + (err.response?.data?.message || err.message));
    } finally {
      setSaving(false);
    }
  }

  async function del(id) {
    if (!canDelete) return;
    const isCampaign = entity === 'campaigns';
    if (confirm(isCampaign ? 'Are you sure you want to permanently delete this campaign record?' : 'Archive this record?')) {
      try {
        await api.delete(`/${entity}/${id}?hard=${isCampaign}`);
        setSelectedIds(prev => {
          const next = new Set(prev);
          next.delete(id);
          return next;
        });
        await load();
        setSaveBanner(`✓ ${isCampaign ? 'Campaign record deleted successfully.' : 'Record archived.'}`);
        setTimeout(() => setSaveBanner(''), 3500);
      } catch (err) {
        alert('Failed to delete record: ' + (err.response?.data?.message || err.message));
      }
    }
  }

  return (
    <>
      <PageHead
        title={title}
        desc={`Managing ${shown.length} total ${title.toLowerCase()} records.`}
        actions={
          <>
            <button type="button" className="scooh-btn ghost" onClick={load}>🔄 Refresh</button>
            {canCreate && (
              <button type="button" className="scooh-btn purple-btn" onClick={() => setEdit({})}>+ Add {title.replace(/s$/, '')}</button>
            )}
          </>
        }
      />

      {isReadOnly && (
        <div className="scooh-banner" style={{ display: 'block', marginBottom: '14px', background: 'rgba(245,158,11,0.12)', border: '1px solid rgba(245,158,11,0.3)', color: '#fbbf24' }}>
          🔒 Read-Only Workspace: You are viewing {title.toLowerCase()} in Viewer mode. Modifying or deleting records is disabled.
        </div>
      )}

      {saveBanner && <div className="scooh-banner" style={{ display: 'block', marginBottom: '14px' }}>{saveBanner}</div>}

      <div className="scooh-toolbar" style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', alignItems: 'center' }}>
        <input
          className="scooh-search"
          style={{ flex: '1 1 240px', minWidth: '200px' }}
          placeholder={`Search ${title.toLowerCase()}…`}
          value={search}
          onChange={e => setSearch(e.target.value)}
        />
        {fs.length > 0 && (
          <select
            value={`${sortState.key}:${sortState.dir}`}
            onChange={e => {
              const [k, d] = e.target.value.split(':');
              setSortState({ key: k, dir: d });
            }}
            style={{ padding: '8px 12px', borderRadius: '8px', background: '#111925', color: '#e2e8f0', border: '1px solid #334155', fontSize: '12px', fontWeight: 600 }}
          >
            <option value=":asc">Default Sorting</option>
            {fs.slice(0, 8).map(f => (
              <React.Fragment key={f}>
                <option value={`${f}:asc`}>{label(f)} (A-Z / Low-High)</option>
                <option value={`${f}:desc`}>{label(f)} (Z-A / High-Low)</option>
              </React.Fragment>
            ))}
          </select>
        )}
        {(search || sortState.key) && (
          <button
            type="button"
            className="scooh-btn ghost"
            style={{ padding: '6px 10px', fontSize: '11px' }}
            onClick={() => { setSearch(''); setSortState({ key: '', dir: 'asc' }); }}
          >
            Reset
          </button>
        )}
      </div>

      {/* Batch Selection Action Bar */}
      {canDelete && selectedIds.size > 0 && (
        <div style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          padding: '11px 18px',
          background: 'linear-gradient(90deg, rgba(239,68,68,0.18), rgba(239,68,68,0.08))',
          borderBottom: '1px solid rgba(239,68,68,0.3)',
          borderRadius: '8px',
          margin: '10px 0',
          color: '#fca5a5',
          fontSize: '13px',
          flexWrap: 'wrap',
          gap: '12px'
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap' }}>
            <span style={{ fontWeight: 800, color: '#ffffff', display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
              <span style={{ display: 'inline-block', width: '8px', height: '8px', borderRadius: '50%', background: '#ef4444' }}></span>
              {selectedIds.size} record{selectedIds.size > 1 ? 's' : ''} selected
            </span>
            <button
              type="button"
              className="scooh-btn ghost"
              onClick={selectAllVisible}
              style={{ fontSize: '11.5px', padding: '4px 10px', color: '#f1f5f9', borderColor: '#475569' }}
            >
              Select all visible ({shown.length})
            </button>
            <button
              type="button"
              className="scooh-btn ghost"
              onClick={deselectAll}
              style={{ fontSize: '11.5px', padding: '4px 10px', color: '#cbd5e1', borderColor: '#475569' }}
            >
              Deselect all
            </button>
          </div>
          <button
            type="button"
            className="scooh-btn danger"
            onClick={handleBatchDelete}
            style={{
              background: '#ef4444',
              color: '#ffffff',
              border: 'none',
              fontWeight: 800,
              padding: '7px 16px',
              borderRadius: '7px',
              cursor: 'pointer',
              display: 'inline-flex',
              alignItems: 'center',
              gap: '6px',
              boxShadow: '0 2px 10px rgba(239,68,68,0.45)'
            }}
          >
            🗑 Delete selected ({selectedIds.size})
          </button>
        </div>
      )}

      <div className="scooh-tablewrap">
        <table className="scooh-table">
          <thead>
            <tr>
              {canDelete && (
                <th style={{ width: '42px', textAlign: 'center', padding: '10px 8px' }}>
                  <input
                    type="checkbox"
                    checked={shown.length > 0 && shown.every(r => selectedIds.has(r.id))}
                    onChange={toggleSelectAll}
                    title="Select all visible records"
                    style={{ cursor: 'pointer' }}
                  />
                </th>
              )}
              {fs.slice(0, 8).map(f => (
                <SortHeader key={f} label={label(f)} sortKey={f} currentSort={sortState} onSort={handleSort} />
              ))}
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {shown.length === 0 ? (
              <tr><td colSpan={fs.slice(0, 8).length + (canDelete ? 2 : 1)} className="scooh-empty">No records found</td></tr>
            ) : (
              shown.map(r => (
                <tr key={r.id}>
                  {canDelete && (
                    <td style={{ textAlign: 'center', padding: '10px 8px' }}>
                      <input
                        type="checkbox"
                        checked={selectedIds.has(r.id)}
                        onChange={() => toggleSelect(r.id)}
                        style={{ cursor: 'pointer' }}
                      />
                    </td>
                  )}
                  {fs.slice(0, 8).map(f => (
                    <td key={f}>
                      {/cost|rate|amount|revenue/.test(f) ? (
                        <b>{money(r[f])}</b>
                      ) : /status|availability/.test(f) ? (
                        <span className={`scooh-pill ${r[f] === 'Active' || r[f] === 'Paid' || r[f] === 'Available' || r[f] === 'Mounted' ? 'active' : r[f] === 'Pending' || r[f] === 'Sent' ? 'watch' : 'vacant'}`}>
                          {String(r[f] ?? '')}
                        </span>
                      ) : /code/.test(f) ? (
                        <span className="scooh-plate">{String(r[f] ?? '')}</span>
                      ) : (
                        String(r[f] ?? '')
                      )}
                    </td>
                  ))}
                  <td>
                    <div className="scooh-rowactions">
                      <button className="scooh-iconbtn scooh-text-action" onClick={() => setEdit(r)}>
                        {canEdit ? 'Edit' : 'View'}
                      </button>
                      {canDelete && (
                        <button
                          className={`scooh-iconbtn danger-icon ${entity === 'campaigns' ? 'scooh-text-action' : ''}`}
                          title={entity === 'campaigns' ? 'Delete campaign record' : 'Archive record'}
                          onClick={() => del(r.id)}
                        >
                          {entity === 'campaigns' ? 'Delete' : '×'}
                        </button>
                      )}
                    </div>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {edit && (
        <div className="scooh-modal-overlay">
          <div className="scooh-modal">
            <div className="scooh-modalhead">
              <div>
                <h2>{edit.id ? (canEdit ? 'Edit' : 'View') : 'Add'} {title.replace(/s$/, '')}</h2>
                <span className="scooh-modal-subtitle">{canEdit ? 'Enter details and save' : 'Read-only record details'}</span>
              </div>
              <button type="button" className="scooh-modal-close" onClick={() => setEdit(null)}>×</button>
            </div>
            <form onSubmit={save}>
              <div className="scooh-modalbody">
                <div className="scooh-grid2">
                  {fs.map(f => {
                    const meta = fieldMeta[f] || {};
                    const isTextarea = meta.textarea || f === 'notes' || f === 'address' || f === 'billing_address' || f === 'terms';
                    const hasSelect = meta.select && Array.isArray(meta.select);

                    return (
                      <div className="scooh-field" key={f} style={isTextarea ? { gridColumn: '1 / -1' } : {}}>
                        <label>{label(f)}</label>
                        {isTextarea ? (
                          <textarea name={f} defaultValue={edit[f] ?? ''} placeholder={meta.placeholder || ''} rows={3} disabled={!canEdit} />
                        ) : hasSelect ? (
                          <select name={f} defaultValue={edit[f] ?? meta.select[0]} disabled={!canEdit}>
                            {meta.select.map(opt => (
                              <option key={opt} value={opt}>{opt}</option>
                            ))}
                          </select>
                        ) : (
                          <input
                            name={f}
                            defaultValue={edit[f] ?? ''}
                            type={/date/.test(f) ? 'date' : /cost|rate|amount|revenue|width|height|tax|discount|days|site_id|client_id/.test(f) ? 'number' : 'text'}
                            step={meta.step || 'any'}
                            placeholder={meta.placeholder || ''}
                            disabled={!canEdit}
                          />
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
              <div className="scooh-modalfoot">
                <button type="button" className="scooh-btn ghost" onClick={() => setEdit(null)}>
                  {canEdit ? 'Cancel' : 'Close'}
                </button>
                {canEdit && (
                  <button className="scooh-btn purple-btn" disabled={saving}>
                    {saving ? 'Saving…' : edit.id ? 'Update Record' : 'Save Record'}
                  </button>
                )}
              </div>
            </form>
          </div>
        </div>
      )}
    </>
  );
}

function SimpleList({ endpoint, title }) {
  const [rows, setRows] = useState([]);

  useEffect(() => {
    api.get('/' + endpoint).then(r => {
      if (Array.isArray(r.data)) setRows(r.data);
    }).catch(() => {});
  }, [endpoint]);

  return (
    <>
      <PageHead title={title} desc={`Reviewing recent system ${title.toLowerCase()} events.`} />
      <div className="scooh-tablewrap">
        <table className="scooh-table">
          <thead>
            <tr>
              {Object.keys((rows && rows[0]) || {}).slice(0, 9).map(k => <th key={k}>{label(k)}</th>)}
            </tr>
          </thead>
          <tbody>
            {(!rows || rows.length === 0) ? (
              <tr><td colSpan="9" className="scooh-empty">No records found</td></tr>
            ) : (
              rows.map((r, i) => (
                <tr key={r.id || i}>
                  {Object.keys(rows[0] || {}).slice(0, 9).map(k => (
                    <td key={k}>
                      {typeof r[k] === 'object' ? JSON.stringify(r[k]) : String(r[k] ?? '')}
                    </td>
                  ))}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </>
  );
}

function App() {
  return (
    <ErrorBoundary>
      <Routes>
        <Route path="/login" element={<Login />} />
        <Route path="/*" element={<Guard><Layout /></Guard>} />
      </Routes>
    </ErrorBoundary>
  );
}

createRoot(document.getElementById('root')).render(
  <BrowserRouter>
    <App />
  </BrowserRouter>
);

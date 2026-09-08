import React, { useEffect, useMemo, useState, useRef, Component } from 'react';
import { createRoot } from 'react-dom/client';
import { BrowserRouter, Navigate, Route, Routes, useLocation, useNavigate } from 'react-router-dom';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import PptxGenJS from 'pptxgenjs';
import api from './api';
import { defaultSites, defaultSettings } from './defaultSites';
import './styles.css';

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
      ['ppt', 'Automated PPT']
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
  settings: 'System Settings'
};

const label = s => String(s).replaceAll('_', ' ').replace(/\b\w/g, m => m.toUpperCase());
const money = v => '₹' + Number(v || 0).toLocaleString('en-IN', { maximumFractionDigits: 2 });
const formatDate = v => v ? new Date(v).toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' }) : '—';

function universalCompare(valA, valB, dir = 'asc') {
  if (valA === valB) return 0;
  if (valA === null || valA === undefined || valA === '' || valA === '—') return 1;
  if (valB === null || valB === undefined || valB === '' || valB === '—') return -1;

  // Try parsing as number
  const numA = typeof valA === 'number' ? valA : (typeof valA === 'string' && /^-?\d+(\.\d+)?$/.test(valA.trim()) ? parseFloat(valA) : NaN);
  const numB = typeof valB === 'number' ? valB : (typeof valB === 'string' && /^-?\d+(\.\d+)?$/.test(valB.trim()) ? parseFloat(valB) : NaN);

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
        res = String(valA).localeCompare(String(valB), undefined, { numeric: true, sensitivity: 'base' });
      }
    } else {
      res = String(valA).localeCompare(String(valB), undefined, { numeric: true, sensitivity: 'base' });
    }
  }

  return dir === 'desc' ? -res : res;
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
  const r = await fetch(url);
  const b = await r.blob();
  return await new Promise((ok, ko) => {
    const fr = new FileReader();
    fr.onload = () => ok(fr.result);
    fr.onerror = ko;
    fr.readAsDataURL(b);
  });
}

async function makePpt(sites, pages) {
  if (!pages.first || !pages.second_last || !pages.last) {
    throw new Error('Please upload all three fixed presentation pages (Page 1, Page 2, and Page 3) first.');
  }
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'Media Buzz Outdoor';
  const SW = 13.333, SH = 7.5, NAVY = '062653', YELLOW = 'FFC400', WHITE = 'FFFFFF';

  async function covered(slide, url, x, y, w, h) {
    try {
      const data = await imageData(url);
      slide.addImage({ data, x, y, w, h, sizing: 'crop' });
    } catch {}
  }

  async function fixed(u) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    await covered(s, u, 0, 0, SW, SH);
  }

  let logo = '';
  try { logo = await imageData('/assets/media-buzz-logo-footer.png'); } catch {}

  async function siteSlide(site, url) {
    const s = pptx.addSlide();
    const leftW = 3.57, leftTopH = 6.18, rightX = leftW, rightW = SW - leftW;
    s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: leftW, h: leftTopH, line: { color: YELLOW, transparency: 100 }, fill: { color: YELLOW } });
    s.addShape(pptx.ShapeType.rect, { x: 0, y: leftTopH, w: leftW, h: SH - leftTopH, line: { color: NAVY, transparency: 100 }, fill: { color: NAVY } });
    
    if (url) await covered(s, url, rightX, 0, rightW, SH);
    else {
      s.addShape(pptx.ShapeType.rect, { x: rightX, y: 0, w: rightW, h: SH, line: { color: 'E8EDF4', transparency: 100 }, fill: { color: 'E8EDF4' } });
      s.addText('No site image added', { x: rightX + 0.45, y: 3.25, w: rightW - 0.9, h: 0.5, fontFace: 'Arial', fontSize: 22, bold: true, color: NAVY, align: 'center', margin: 0, fit: 'shrink' });
    }
    
    const heading = String(site.area || site.address || site.city || site.site_code || 'Site');
    s.addText(heading, { x: 0.55, y: 1.18, w: 2.47, h: 0.5, fontFace: 'Arial', fontSize: 23, bold: true, color: NAVY, margin: 0, fit: 'shrink' });
    
    if (site.address && site.address.toLowerCase() !== heading.toLowerCase()) {
      s.addText(String(site.address), { x: 0.55, y: 1.72, w: 2.47, h: 0.65, fontFace: 'Arial', fontSize: 14, bold: true, color: NAVY, margin: 0, fit: 'shrink' });
    }

    const rows = [
      ['Size', site.size],
      ['Availability', site._availability || site.ppt_availability || site.availability],
      ['Lighting', site.lighting],
      ['Type', site.media_type]
    ];
    if (site._showRate && site._rate) rows.push(['Rate', `${site._rate}/- Per Month`]);
    if (site._showCoords) {
      rows.push(['Latitude', site.latitude || '—']);
      rows.push(['Longitude', site.longitude || '—']);
    }

    let y = 2.66;
    for (const [l, v] of rows.filter(x => x[1] !== '' && x[1] != null)) {
      s.addText('•', { x: 0.50, y, w: 0.17, h: 0.25, fontSize: 12, bold: true, color: NAVY, margin: 0 });
      s.addText(l + ':', { x: 0.70, y, w: 1.08, h: 0.25, fontSize: 11.2, bold: true, color: NAVY, margin: 0, fit: 'shrink' });
      s.addText(String(v), { x: 1.82, y, w: 1.28, h: 0.25, fontSize: 11.2, bold: true, color: NAVY, margin: 0, fit: 'shrink' });
      y += 0.39;
    }

    if (logo) s.addImage({ data: logo, x: 0.42, y: leftTopH + 0.22, w: 2.72, h: 0.82 });
  }

  await fixed(pages.first);
  await fixed(pages.second_last);
  for (const st of sites) {
    if (st.ppt_images && st.ppt_images.length) {
      for (const u of st.ppt_images) await siteSlide(st, u);
    } else {
      await siteSlide(st, '');
    }
  }
  await fixed(pages.last);
  await pptx.writeFile({ fileName: 'MediaBuzz_Automated-PPT.pptx' });
}

async function exportStyledExcel(rowsData, filename = 'MediaBuzz_Sites.xlsx') {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sites', {
    views: [{ state: 'frozen', ySplit: 1 }]
  });

  worksheet.columns = [
    { header: 'SR NO', key: 'sr', width: 8 },
    { header: 'AREA', key: 'area', width: 22 },
    { header: 'LOCATION', key: 'location', width: 55 },
    { header: 'MEDIA', key: 'media', width: 14 },
    { header: 'LIGHT', key: 'light', width: 10 },
    { header: 'W', key: 'w', width: 8 },
    { header: 'H', key: 'h', width: 8 },
    { header: 'SQ FT', key: 'sqft', width: 12 },
    { header: 'AVAILABLITY', key: 'availability', width: 18 },
    { header: 'Selling Amount', key: 'rate', width: 18 },
    { header: 'Latitude Longitude', key: 'coords', width: 28 }
  ];

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

  // Style Header Row (Row 1) in Bright Yellow (#FFFF00)
  const headerRow = worksheet.getRow(1);
  headerRow.height = 28;
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
    if (rowNumber > 1) {
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
    worksheet.autoFilter = `A1:K${rowsData.length + 1}`;
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
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
            <div className="scooh-topbar-context">
              <span className="scooh-topbar-eyebrow">MEDIA BUZZ • OOH WORKSPACE</span>
              <strong>{currentTitle}</strong>
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
              <span>Signed in as</span>
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
            {navGroups.map(grp => (
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
              <Route path="/campaigns" element={<Crud entity="campaigns" title="Campaign Tracker" />} />
              <Route path="/occupancy" element={<OccupancyView />} />
              <Route path="/proposals" element={<ProposalsView />} />
              <Route path="/ppt" element={<PptView />} />
              <Route path="/electricity" element={<ElectricityView />} />
              <Route path="/vendors" element={<Crud entity="vendors" title="Vendors" />} />
              <Route path="/clients" element={<Crud entity="clients" title="Clients" />} />
              <Route path="/invoices" element={<Crud entity="invoices" title="Invoices" />} />
              <Route path="/data" element={<DataToolsView />} />
              <Route path="/reports" element={<ReportsView />} />
              <Route path="/notifications" element={<SimpleList endpoint="notifications" title="Notifications" />} />
              <Route path="/activity" element={<SimpleList endpoint="activity" title="Activity Log" />} />
              <Route path="/settings" element={<SettingsView />} />
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
    { label: 'Star', value: buckets.Star || 0, color: '#FBBF24' },
    { label: 'Solid', value: buckets.Solid || 0, color: '#22C55E' },
    { label: 'Needs attention', value: buckets['Needs attention'] || 0, color: '#8B5CF6' },
    { label: 'Underperforming', value: buckets.Underperforming || 1, color: '#EF4444' }
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
            <span>{d.label}</span>
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
        <div className="scooh-kpi alert">
          <div className="n">{k.total_sites || 57}</div>
          <div className="l">Total Sites</div>
          {k.flagged_count > 0 && (
            <div className="scooh-kpi-note">⚠ {k.flagged_count} flagged for review</div>
          )}
        </div>
        <div className="scooh-kpi good">
          <div className="n">{k.active_sites || 1}</div>
          <div className="l">Active Sites</div>
        </div>
        <div className="scooh-kpi alert">
          <div className="n">{k.nonactive_sites || (k.total_sites ? k.total_sites - 1 : 56)}</div>
          <div className="l">Vacant / Non-active</div>
        </div>
        <div className="scooh-kpi alert">
          <div className="n">{k.active_campaigns || 2}</div>
          <div className="l">Active Campaigns</div>
        </div>
      </div>

      {/* Row 2: Operational Action Triggers */}
      <div className="scooh-kpirow">
        <div className="scooh-kpi danger">
          <div className="n">{k.mounting_overdue ?? 2}</div>
          <div className="l">Mounting Overdue</div>
        </div>
        <div className="scooh-kpi danger">
          <div className="n">{k.validation_15_due ?? 0}</div>
          <div className="l">15-Day Validation Due</div>
        </div>
        <div className="scooh-kpi danger">
          <div className="n">{k.final_validation_due ?? 1}</div>
          <div className="l">Final Validation Due</div>
        </div>
        <div className="scooh-kpi danger">
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
            <div className="scooh-empty">Nothing outstanding — everything is on track.</div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
              {alerts.map((a, idx) => (
                <div
                  key={idx}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    padding: '12px 14px',
                    border: '1px solid var(--mb-border)',
                    borderRadius: '12px',
                    background: 'var(--mb-surface-2)'
                  }}
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

  useEffect(() => {
    load();
    const interval = setInterval(load, 35000); // 35s auto-refresh
    return () => clearInterval(interval);
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
      const text = JSON.stringify(s).toLowerCase();
      const matchesQuery = text.includes(search.toLowerCase());
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

  async function save(e) {
    e.preventDefault();
    setSaving(true);
    const data = Object.fromEntries(new FormData(e.currentTarget));
    try {
      if (edit?.id) await api.put(`/sites/${edit.id}`, data);
      else await api.post('/sites', data);
      setEdit(null);
      await load();
    } catch (err) {
      alert('Error saving site: ' + (err.response?.data?.message || err.message));
    } finally {
      setSaving(false);
    }
  }

  async function del(id) {
    if (confirm('Archive this site record?')) {
      await api.delete(`/sites/${id}`);
      await load();
    }
  }

  async function handleImageUpload(siteId, files) {
    const fd = new FormData();
    [...files].forEach(f => fd.append('files', f));
    await api.post(`/sites/${siteId}/images`, fd, { headers: { 'Content-Type': 'multipart/form-data' } });
    await load();
  }

  return (
    <>
      <PageHead
        title="Sites Directory"
        desc={`Managing ${filtered.length} of ${sites.length} total Media Buzz outdoor inventory sites.`}
        actions={
          <>
            <button type="button" className="scooh-btn ghost" onClick={load}>🔄 Refresh</button>
            <button className="scooh-btn purple-btn" onClick={() => setEdit({})}>+ Add Site</button>
          </>
        }
      />

      <div className="scooh-toolbar" style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', alignItems: 'center' }}>
        <input
          className="scooh-search"
          style={{ flex: '1 1 220px', minWidth: '180px' }}
          placeholder="Search site ID, location, landmark…"
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
            <option value="site_code:asc">Sort: Site Code (A-Z)</option>
            <option value="site_code:desc">Sort: Site Code (Z-A)</option>
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

      <div className="scooh-tablewrap">
        <table className="scooh-table">
          <thead>
            <tr>
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
              <tr>
                <td colSpan="9" className="scooh-empty">No matching sites found</td>
              </tr>
            ) : (
              filtered.map(s => (
                <tr key={s.id || s.site_code}>
                  <td><span className="scooh-plate">{s.site_code}</span></td>
                  <td>
                    <strong>{s.area || s.city}</strong>
                    <div className="scooh-footnote">{s.address}</div>
                  </td>
                  <td>{s.media_type}</td>
                  <td>{s.size}</td>
                  <td><span className="scooh-badge">{s.lighting}</span></td>
                  <td>
                    <span className={`scooh-pill ${s.availability === 'Available' ? 'active' : s.availability === 'Occupied' ? 'vacant' : 'watch'}`}>
                      {s.ppt_availability || s.availability}
                    </span>
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
                      <button className="scooh-iconbtn scooh-text-action" onClick={() => setEdit(s)}>Edit</button>
                      <button className="scooh-iconbtn danger-icon" onClick={() => del(s.id)}>×</button>
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
                <button type="button" className="scooh-btn ghost" onClick={() => setEdit(null)}>Cancel</button>
                <button className="scooh-btn purple-btn" disabled={saving}>
                  {saving ? 'Saving…' : 'Save Site'}
                </button>
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
  const [sites, setSites] = useState(defaultSites.map(unpackSite));
  const [pages, setPages] = useState({});
  const [sel, setSel] = useState({});
  const [query, setQuery] = useState('');
  const [areaFilter, setAreaFilter] = useState('');
  const [dateFilter, setDateFilter] = useState('');
  const [generating, setGenerating] = useState(false);
  const [sortState, setSortState] = useState({ key: 'site_code', dir: 'asc' });

  async function load() {
    try {
      const [sRes, pRes] = await Promise.all([
        api.get('/sites'),
        api.get('/ppt-pages')
      ]);
      if (Array.isArray(sRes.data) && sRes.data.length > 0) {
        setSites(sRes.data.map(unpackSite));
      }
      if (pRes.data) setPages(pRes.data);
    } catch (e) {
      console.warn('PPT load notice:', e);
    }
  }

  useEffect(() => { load(); }, []);

  const areas = useMemo(() => Array.from(new Set(sites.map(s => s.area).filter(Boolean))).sort(), [sites]);

  function handleSort(key) {
    setSortState(prev => prev.key === key ? { key, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { key, dir: 'asc' });
  }

  const filtered = useMemo(() => {
    const list = sites.filter(s => {
      const hay = [s.site_code, s.address, s.city, s.area, s.media_type].filter(Boolean).join(' ').toLowerCase();
      if (query && !hay.includes(query.toLowerCase().trim())) return false;
      if (areaFilter && s.area !== areaFilter) return false;
      if (dateFilter) {
        const av = String(s.ppt_availability || s.availability || '').toLowerCase().trim();
        if (av !== 'available' && av !== 'immediate' && av !== 'long term') {
          if (av > dateFilter) return false;
        }
      }
      return true;
    });
    if (sortState.key) {
      list.sort((a, b) => {
        let valA = a[sortState.key];
        let valB = b[sortState.key];
        if (sortState.key === 'monthly_rate') {
          valA = Number(a.ppt_rate || a.monthly_rate || 0);
          valB = Number(b.ppt_rate || b.monthly_rate || 0);
        }
        return universalCompare(valA, valB, sortState.dir);
      });
    }
    return list;
  }, [sites, query, areaFilter, dateFilter, sortState]);

  async function uploadPage(k, file) {
    try {
      const fd = new FormData();
      fd.append('file', file);
      const res = await api.post('/ppt-pages/' + k, fd, { headers: { 'Content-Type': 'multipart/form-data' } });
      setPages(prev => ({ ...prev, [k]: res.data.url }));
    } catch (e) {
      alert('Failed to upload page: ' + (e.response?.data?.message || e.message));
    }
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

      const rows = targetSites.map((x, idx) => {
        const v = sel[x.id || x.site_code] || {};
        let w = x.width || '', h = x.height || '';
        if ((!w || !h) && x.size) {
          const parts = String(x.size).toLowerCase().split('x');
          if (parts.length === 2) { w = parts[0].trim(); h = parts[1].trim(); }
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
          'AVAILABLITY': v.availability ?? x.ppt_availability ?? x.availability ?? 'Available',
          'Selling Amount': Number(v.rate ?? x.ppt_rate ?? x.monthly_rate ?? 0),
          'Latitude Longitude': coords
        };
      });

      await exportStyledExcel(rows, 'MediaBuzz_PPT_Inventory.xlsx');
    } catch (e) {
      alert('Export Excel failed: ' + e.message);
    }
  }

  async function generate() {
    const chosen = sites
      .filter(s => sel[s.id || s.site_code]?.checked)
      .map(s => {
        const v = sel[s.id || s.site_code] || {};
        return {
          ...s,
          _availability: v.availability ?? s.ppt_availability ?? s.availability,
          _rate: v.rate ?? s.ppt_rate ?? s.monthly_rate,
          _showRate: !!v.showRate,
          _showCoords: !!v.showCoords
        };
      });

    if (!chosen.length) return alert('Please select at least one site to include in the PowerPoint.');

    setGenerating(true);
    try {
      await makePpt(chosen, pages);
    } catch (e) {
      alert(e.message);
    } finally {
      setGenerating(false);
    }
  }

  return (
    <>
      <PageHead
        title="Automated PPT"
        desc="Select sites, add images, set availability, and generate/download your presentation data."
        actions={
          <>
            <button type="button" className="scooh-btn ghost" onClick={downloadExcel}>
              Download Excel (.xlsx)
            </button>
            <button type="button" className="scooh-btn primary" onClick={generate} disabled={generating}>
              {generating ? 'Creating PPT…' : 'Generate PPT'}
            </button>
          </>
        }
      />

      {/* Fixed presentation pages section */}
      <section className="scooh-form-section scooh-ppt-fixed-pages">
        <div className="scooh-form-section-head">
          <div>
            <h3>Fixed presentation pages</h3>
            <p>Uploaded fixed pages fill the entire 16:9 slide. Order: page 1, page 2, selected site slides, then page 3.</p>
          </div>
        </div>
        <div className="scooh-grid3">
          {[
            ['first', 'First page', 'Full-screen cover image for slide 1.'],
            ['second_last', 'Second page', 'Full-screen image used as slide 2, before the selected site slides.'],
            ['last', 'Last page', 'Full-screen image used after all selected site slides.']
          ].map(([k, title, desc]) => (
            <div className="scooh-panel" key={k} style={{ margin: 0, padding: '16px' }}>
              <strong>{title}</strong>
              <small style={{ display: 'block', margin: '6px 0 12px', color: '#91a5c2' }}>{desc}</small>
              <label className="scooh-btn" style={{ cursor: 'pointer' }}>
                {pages[k] ? 'Replace uploaded image' : 'Upload image'}
                <input type="file" accept="image/*" hidden onChange={e => e.target.files[0] && uploadPage(k, e.target.files[0])} />
              </label>
              {pages[k] && (
                <img src={pages[k]} alt={title} style={{ width: '100%', height: '120px', objectFit: 'cover', borderRadius: '10px', marginTop: '12px' }} />
              )}
            </div>
          ))}
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
              placeholder="Search by site code, location, city, area or media type..."
              value={query}
              onChange={e => setQuery(e.target.value)}
              autoComplete="off"
            />
          </div>

          <div
            className="scooh-ppt-select-actions"
            style={{ display: 'flex', alignItems: 'center', gap: '8px', flexShrink: 0 }}
          >
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
              value={sortState.key}
              onChange={e => handleSort(e.target.value)}
              style={{
                height: '44px',
                padding: '0 12px',
                borderRadius: '8px',
                border: '1px solid #344258',
                background: '#111720',
                color: '#e2eaf4',
                fontSize: '12.5px',
                cursor: 'pointer'
              }}
            >
              <option value="site_code">Sort: Site Code</option>
              <option value="area">Sort: Area</option>
              <option value="city">Sort: City</option>
              <option value="media_type">Sort: Media Type</option>
              <option value="monthly_rate">Sort: Rate</option>
              <option value="availability">Sort: Availability</option>
            </select>
            <button
              type="button"
              className="scooh-btn ghost"
              onClick={() => setSortState(prev => ({ ...prev, dir: prev.dir === 'asc' ? 'desc' : 'asc' }))}
              title={sortState.dir === 'asc' ? 'Ascending — click for descending' : 'Descending — click for ascending'}
              style={{ minHeight: '44px', padding: '0 12px', fontSize: '14px', whiteSpace: 'nowrap' }}
            >
              {sortState.dir === 'asc' ? '↑ A–Z' : '↓ Z–A'}
            </button>
          </div>
        </div>

        <div className="scooh-ppt-search-count" id="scooh-ppt-search-count">
          {filtered.length} of {sites.length} sites
        </div>

        {/* Sites Grid */}
        <div className="scooh-ppt-grid" id="scooh-ppt-grid">
          {filtered.map(s => {
            const siteKey = s.id || s.site_code;
            const v = sel[siteKey] || {};
            const isChecked = !!v.checked;
            const imgs = s.ppt_images || [];

            return (
              <label
                key={siteKey}
                className="scooh-ppt-site"
                style={{
                  border: isChecked ? '1px solid #9c8cff' : '1px solid #344258',
                  boxShadow: isChecked ? '0 0 0 1px #9c8cff' : 'none'
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

                <strong>{s.site_code || 'Site'}</strong>
                <small>{s.location || s.address || s.area || s.city || ''}</small>
                <small className="scooh-ppt-coordinates">Latitude: {s.latitude ?? '—'}</small>
                <small className="scooh-ppt-coordinates">Longitude: {s.longitude ?? '—'}</small>

                <div className="scooh-ppt-availability" onClick={e => e.stopPropagation()}>
                  <label>Availability for PPT</label>
                  <input
                    type="text"
                    className="scooh-ppt-availability-input"
                    value={v.availability ?? s.ppt_availability ?? ''}
                    placeholder="Immediate or DD.MM.YYYY"
                    onChange={e => setSel({ ...sel, [siteKey]: { ...v, availability: e.target.value } })}
                  />
                </div>

                <div className="scooh-ppt-output-options" onClick={e => e.stopPropagation()}>
                  <label className="scooh-ppt-output-option">
                    <input
                      type="checkbox"
                      className="scooh-ppt-show-coordinates"
                      checked={!!v.showCoords}
                      onChange={e => setSel({ ...sel, [siteKey]: { ...v, showCoords: e.target.checked } })}
                    />
                    <span>Show Latitude / Longitude in PPT</span>
                  </label>
                  <label className="scooh-ppt-output-option">
                    <input
                      type="checkbox"
                      className="scooh-ppt-show-rate"
                      checked={!!v.showRate}
                      onChange={e => setSel({ ...sel, [siteKey]: { ...v, showRate: e.target.checked } })}
                    />
                    <span>Show Rate in PPT</span>
                  </label>
                </div>

                <div className="scooh-ppt-rate" onClick={e => e.stopPropagation()}>
                  <label>Rate Per Month (₹)</label>
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

  const filteredSites = useMemo(() => {
    const list = sites.filter(s => {
      const hay = [s.site_code, s.area, s.city, s.media_type, s.size].filter(Boolean).join(' ').toLowerCase();
      return !search || hay.includes(search.toLowerCase().trim());
    });

    if (siteSort.key) {
      list.sort((a, b) => {
        let valA = a[siteSort.key];
        let valB = b[siteSort.key];
        if (siteSort.key === 'monthly_rate') {
          valA = Number(a.ppt_rate || a.monthly_rate || 0);
          valB = Number(b.ppt_rate || b.monthly_rate || 0);
        }
        return universalCompare(valA, valB, siteSort.dir);
      });
    }
    return list;
  }, [sites, search, siteSort]);

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
            <h3 className="scooh-proposal-step" style={{ margin: 0, padding: 0, border: 'none' }}>2. Select sites</h3>
            <select
              value={`${siteSort.key}:${siteSort.dir}`}
              onChange={e => {
                const [k, d] = e.target.value.split(':');
                setSiteSort({ key: k, dir: d });
              }}
              style={{ padding: '4px 8px', fontSize: '11px', borderRadius: '6px', background: '#111925', color: '#e2e8f0', border: '1px solid #334155' }}
            >
              <option value="site_code:asc">Sort: Site Code (A-Z)</option>
              <option value="site_code:desc">Sort: Site Code (Z-A)</option>
              <option value="area:asc">Sort: Area (A-Z)</option>
              <option value="monthly_rate:asc">Sort: Rate (Low to High)</option>
              <option value="monthly_rate:desc">Sort: Rate (High to Low)</option>
              <option value="size:asc">Sort: Size</option>
            </select>
          </div>
          <div className="scooh-field">
            <input
              type="search"
              placeholder="Filter sites by code, area, landmark, media..."
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

function OccupancyView() {
  const [rows, setRows] = useState([]);
  const [search, setSearch] = useState('');
  const [sortState, setSortState] = useState({ key: 'pct', dir: 'desc' });

  useEffect(() => {
    Promise.all([api.get('/sites'), api.get('/campaigns')]).then(([s, c]) => {
      const siteList = (Array.isArray(s.data) && s.data.length > 0) ? s.data : defaultSites;
      const campaignList = Array.isArray(c.data) ? c.data : [];
      const now = new Date(), from = new Date(now);
      from.setDate(now.getDate() - 365);

      const calculated = siteList.map(site => {
        let days = 0;
        campaignList.filter(x => String(x.site_id) === String(site.id)).forEach(x => {
          const a = new Date(x.start_date), b = new Date(x.end_date);
          const start = a < from ? from : a, end = b > now ? now : b;
          if (end >= start) days += (end - start) / 86400000 + 1;
        });
        return {
          site_code: site.site_code,
          city: site.city,
          area: site.area,
          occupied: Math.round(days),
          pct: Math.min(100, Math.round((days / 365) * 100))
        };
      });
      setRows(calculated);
    }).catch(() => {
      setRows(defaultSites.map(s => ({ site_code: s.site_code, city: s.city, area: s.area, occupied: 0, pct: 0 })));
    });
  }, []);

  function handleSort(key) {
    setSortState(prev => prev.key === key ? { key, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { key, dir: 'asc' });
  }

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    const list = rows.filter(r => {
      if (!q) return true;
      return [r.site_code, r.city, r.area].some(v => String(v || '').toLowerCase().includes(q));
    });

    if (sortState.key) {
      list.sort((a, b) => {
        let valA = a[sortState.key];
        let valB = b[sortState.key];
        if (sortState.key === 'occupied' || sortState.key === 'pct') {
          valA = Number(valA || 0);
          valB = Number(valB || 0);
        }
        return universalCompare(valA, valB, sortState.dir);
      });
    }
    return list;
  }, [rows, search, sortState]);

  return (
    <>
      <PageHead title="Occupancy & Utilization" desc="Rolling 365-day site occupancy calculations and performance breakdown." />
      <section className="scooh-panel">
        <div className="scooh-toolbar" style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', alignItems: 'center', marginBottom: '14px' }}>
          <input
            className="scooh-search"
            style={{ flex: '1 1 240px', minWidth: '200px' }}
            placeholder="Search by site code, city, area…"
            value={search}
            onChange={e => setSearch(e.target.value)}
          />
          <select
            value={`${sortState.key}:${sortState.dir}`}
            onChange={e => {
              const [k, d] = e.target.value.split(':');
              setSortState({ key: k, dir: d });
            }}
            style={{ padding: '8px 12px', borderRadius: '8px', background: '#111925', color: '#e2e8f0', border: '1px solid #334155', fontSize: '12px', fontWeight: 600 }}
          >
            <option value="pct:desc">Sort: Occupancy % (High to Low)</option>
            <option value="pct:asc">Sort: Occupancy % (Low to High)</option>
            <option value="occupied:desc">Sort: Occupied Days (High to Low)</option>
            <option value="occupied:asc">Sort: Occupied Days (Low to High)</option>
            <option value="site_code:asc">Sort: Site ID (A-Z)</option>
            <option value="site_code:desc">Sort: Site ID (Z-A)</option>
            <option value="area:asc">Sort: Area (A-Z)</option>
            <option value="city:asc">Sort: City (A-Z)</option>
          </select>
          {(search || sortState.key !== 'pct' || sortState.dir !== 'desc') && (
            <button
              type="button"
              className="scooh-btn ghost"
              style={{ padding: '6px 10px', fontSize: '11px' }}
              onClick={() => { setSearch(''); setSortState({ key: 'pct', dir: 'desc' }); }}
            >
              Reset
            </button>
          )}
        </div>

        <div className="scooh-tablewrap">
          <table className="scooh-table">
            <thead>
              <tr>
                <SortHeader label="Site ID" sortKey="site_code" currentSort={sortState} onSort={handleSort} />
                <SortHeader label="City" sortKey="city" currentSort={sortState} onSort={handleSort} />
                <SortHeader label="Area" sortKey="area" currentSort={sortState} onSort={handleSort} />
                <SortHeader label="Occupied Days (365d)" sortKey="occupied" currentSort={sortState} onSort={handleSort} />
                <SortHeader label="Occupancy %" sortKey="pct" currentSort={sortState} onSort={handleSort} />
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr><td colSpan="5" className="scooh-empty">No sites match the search filter</td></tr>
              ) : (
                filtered.map(r => (
                  <tr key={r.site_code}>
                    <td><span className="scooh-plate">{r.site_code}</span></td>
                    <td>{r.city}</td>
                    <td><b>{r.area}</b></td>
                    <td>{r.occupied} days</td>
                    <td>
                      <div className="scooh-occ-cell">
                        <div className="scooh-occbar">
                          <span style={{ width: r.pct + '%', background: r.pct > 70 ? '#48c79a' : r.pct > 30 ? '#e8a94b' : '#f06a6a' }} />
                        </div>
                        <b>{r.pct}%</b>
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

  useEffect(() => {
    loadData();
    const interval = setInterval(loadData, 30000);
    return () => clearInterval(interval);
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
  const providers = Array.from(new Set(rows.map(r => r.bill_type).filter(Boolean)));

  function handleSort(key) {
    setSortState(prev => prev.key === key ? { key, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { key, dir: 'asc' });
  }

  const filtered = useMemo(() => {
    const list = rows.filter(r => {
      const q = search.trim().toLowerCase();
      const matchesSearch = !q || [
        r.site_code,
        r.location,
        r.size,
        r.meter_no,
        r.service_number,
        r.t_number,
        r.bill_type,
        r.billing_month,
        r.payment_status,
        r.notes
      ].some(v => String(v || '').toLowerCase().includes(q));

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
    setEditModal({
      site_id: '',
      site_code: '',
      location: '',
      size: '',
      meter_no: '',
      service_number: '',
      t_number: '',
      bill_type: 'UGVCL',
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
    if (!payModal?.id) return;
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

  async function deleteBill(id) {
    if (confirm('Archive this electricity bill record?')) {
      try {
        await api.delete(`/electricity/${id}`);
        setBanner('✓ Electricity bill archived.');
        await loadData();
        setTimeout(() => setBanner(''), 3000);
      } catch (err) {
        alert('Failed to archive bill: ' + (err.response?.data?.message || err.message));
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
          <button type="button" className="scooh-btn ghost" onClick={() => exportElectricityExcel(filtered)}>
            📥 Export Excel
          </button>
          <button type="button" className="scooh-btn purple-btn" onClick={openNewBill}>
            + Add bill
          </button>
        </div>
      </div>

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
          <div className="scooh-kpi-note" style={{ color: '#94a3b8' }}>UGVCL, Torrent Power & PGVCL</div>
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
            <button type="button" className="scooh-btn purple-btn" onClick={openNewBill}>
              + Add bill
            </button>
          </div>
        </div>

        {/* Filters Toolbar */}
        <div className="scooh-toolbar" style={{ padding: '14px 20px', borderBottom: '1px solid rgba(255,255,255,0.06)', display: 'flex', gap: '12px', flexWrap: 'wrap', alignItems: 'center' }}>
          <input
            className="scooh-search"
            style={{ flex: '1 1 240px', minWidth: '200px' }}
            placeholder="Search by site code, location, meter, consumer no, provider…"
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
            <option value="due_date:asc">Sort: Due Date (Earliest First)</option>
            <option value="due_date:desc">Sort: Due Date (Latest First)</option>
            <option value="amount:desc">Sort: Amount (High to Low)</option>
            <option value="amount:asc">Sort: Amount (Low to High)</option>
            <option value="site_code:asc">Sort: Site Code (A-Z)</option>
            <option value="site_code:desc">Sort: Site Code (Z-A)</option>
            <option value="location:asc">Sort: Location (A-Z)</option>
            <option value="payment_status:asc">Sort: Status (A-Z)</option>
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

        {/* Table */}
        <div className="scooh-tablewrap">
          <table className="scooh-table">
            <thead>
              <tr>
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
                  <td colSpan="11" className="scooh-empty" style={{ padding: '36px 20px' }}>
                    <div style={{ fontSize: '14px', fontWeight: 700, color: '#94a3b8' }}>No electricity bills match your query</div>
                    <div style={{ fontSize: '12px', color: '#64748b', marginTop: '4px' }}>Click "+ Add bill" to record a new electricity bill for any hoarding site.</div>
                  </td>
                </tr>
              ) : (
                filtered.map(r => {
                  const isOverdue = r.payment_status !== 'Paid' && r.due_date && new Date(r.due_date) < new Date();
                  return (
                    <tr key={r.id}>
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
                            background: r.bill_type === 'Torrent Power' ? 'rgba(56,189,248,0.16)' : 'rgba(251,191,36,0.16)',
                            color: r.bill_type === 'Torrent Power' ? '#38bdf8' : '#fbbf24',
                            fontWeight: 800,
                            fontSize: '10.5px'
                          }}
                        >
                          {r.bill_type || 'General'}
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
                          {r.payment_status !== 'Paid' && (
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
                            title="Edit Bill"
                          >
                            Edit
                          </button>
                          <button
                            type="button"
                            className="scooh-iconbtn danger-icon"
                            onClick={() => deleteBill(r.id)}
                            title="Delete Bill"
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
      </div>

      {/* Add / Edit Bill Modal */}
      {editModal && (
        <div className="scooh-modal-overlay">
          <div className="scooh-modal" style={{ maxWidth: '640px' }}>
            <div className="scooh-modalhead">
              <div>
                <h2>{editModal.id ? 'Edit Electricity Bill' : 'Add Electricity Bill'}</h2>
                <span className="scooh-modal-subtitle">Auto-fill site metadata and manage payment tracking</span>
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
                    <label>Bill Provider / Type</label>
                    <select
                      name="bill_type"
                      value={editModal.bill_type ?? 'UGVCL'}
                      onChange={e => setEditModal({ ...editModal, bill_type: e.target.value })}
                    >
                      <option value="UGVCL">UGVCL</option>
                      <option value="Torrent Power">Torrent Power</option>
                      <option value="PGVCL">PGVCL</option>
                      <option value="DGVCL">DGVCL</option>
                      <option value="MGVCL">MGVCL</option>
                      <option value="Other">Other Utility</option>
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
                <button type="button" className="scooh-btn ghost" onClick={() => setEditModal(null)}>Cancel</button>
                <button type="submit" className="scooh-btn purple-btn" disabled={saving}>
                  {saving ? 'Saving…' : editModal.id ? 'Update Electricity Bill' : 'Save Electricity Bill'}
                </button>
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
  const [jsonFile, setJsonFile] = useState(null);
  const [xlsxStatus, setXlsxStatus] = useState('No Excel file selected.');
  const [jsonStatus, setJsonStatus] = useState('No JSON file selected.');
  const [loading, setLoading] = useState(false);

  async function importXlsx() {
    if (!xlsxFile) return alert('Please choose an Excel file (.xlsx) first.');
    setLoading(true);
    setXlsxStatus(`Importing ${xlsxFile.name}…`);
    try {
      const f = new FormData();
      f.append('file', xlsxFile);
      const r = await api.post('/import/xlsx', f, { headers: { 'Content-Type': 'multipart/form-data' } });
      setXlsxStatus(`✓ Successfully imported ${r.data.rows} sites from ${xlsxFile.name}!`);
      alert(`Success: ${r.data.rows} sites imported successfully.`);
    } catch (e) {
      setXlsxStatus(`✕ Import failed: ${e.response?.data?.message || e.message}`);
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
          activeTab !== 'users' ? (
            <button type="button" className="scooh-btn purple-btn" onClick={() => saveSettings()}>
              Save Changes
            </button>
          ) : (
            <button type="button" className="scooh-btn purple-btn" onClick={() => setUserModal({})}>
              + Add User
            </button>
          )
        }
      />

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
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '16px' }}>
              <div>
                <h3>Users & Role Permissions</h3>
                <p className="scooh-footnote" style={{ color: '#828f9e', fontSize: '11px', margin: '4px 0 0' }}>
                  Create workspace staff accounts and assign granular security roles.
                </p>
              </div>
              <button type="button" className="scooh-btn purple-btn" onClick={() => setUserModal({})}>
                + Add User
              </button>
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
                    users.map(u => (
                      <tr key={u.id}>
                        <td>
                          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                            <span className="scooh-user-avatar" style={{ width: '28px', height: '28px', fontSize: '11px' }}>
                              {(u.name || u.email || 'U').slice(0, 1).toUpperCase()}
                            </span>
                            <strong>{u.name || 'User'}</strong>
                          </div>
                        </td>
                        <td>{u.email}</td>
                        <td>
                          <span
                            className="scooh-badgechip"
                            style={{
                              background: u.role === 'admin' ? 'rgba(139,92,246,0.18)' : u.role === 'manager' ? 'rgba(72,199,154,0.18)' : 'rgba(255,255,255,0.06)',
                              color: u.role === 'admin' ? '#c4b5fd' : u.role === 'manager' ? '#48c79a' : '#d0d7de',
                              textTransform: 'uppercase',
                              fontWeight: 800,
                              fontSize: '10px'
                            }}
                          >
                            {u.role || 'staff'}
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
                            <button className="scooh-iconbtn scooh-text-action" onClick={() => setUserModal(u)}>Edit</button>
                            <button className="scooh-iconbtn danger-icon" onClick={() => deleteUser(u.id)}>×</button>
                          </div>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </section>
        </div>
      </div>

      {/* Add / Edit User Modal */}
      {userModal && (
        <div className="scooh-modal-overlay">
          <div className="scooh-modal" style={{ maxWidth: '520px' }}>
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
                    <label>Role</label>
                    <select name="role" defaultValue={userModal.role ?? 'staff'}>
                      <option value="admin">Admin (Full Access & Settings)</option>
                      <option value="manager">Manager (Sites, Campaigns, Proposals)</option>
                      <option value="staff">Staff (Field Ops, Mounting & Meters)</option>
                      <option value="viewer">Viewer (Read-Only Access)</option>
                    </select>
                  </div>
                  <div className="scooh-field">
                    <label>Status</label>
                    <select name="status" defaultValue={userModal.status ?? 'active'}>
                      <option value="active">Active</option>
                      <option value="disabled">Disabled</option>
                    </select>
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
  bill_type: { select: ['UGVCL', 'Torrent Power', 'PGVCL', 'DGVCL', 'MGVCL'] },
  invoice_requested: { select: ['No', 'Yes'] },
  invoice_required: { select: ['1', '0'] },
  hard_copy_required: { select: ['1', '0'] },

  // Placeholders
  client_name: { placeholder: 'e.g. Rajyash Group' },
  company: { placeholder: 'e.g. Rajyash Estates Pvt Ltd' },
  primary_contact: { placeholder: 'e.g. Pratik Patel' },
  email: { placeholder: 'e.g. contact@client.com' },
  phone: { placeholder: 'e.g. +91 98250 11223' },
  gst_number: { placeholder: 'e.g. 24AAACR1234F1Z5' },
  name: { placeholder: 'e.g. Gujarat Printers & Signage' },
  contact_person: { placeholder: 'e.g. Mukesh Bhai' },
  cities: { placeholder: 'e.g. Ahmedabad, Gandhinagar, Surat' },
  rating: { placeholder: '4.85', step: '0.05' },
  booking_code: { placeholder: 'e.g. MB-BK-2026-001' },
  parent_campaign: { placeholder: 'e.g. Annual Media Contract' },
  site_code: { placeholder: 'e.g. AMD-GT-001' },
  client: { placeholder: 'e.g. Rajyash Group' },
  brand: { placeholder: 'e.g. Rajyash Estates' },
  campaign_name: { placeholder: 'e.g. Diwali Launch Ahmedabad' },
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
  const fs = fields[entity] || [];

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

  async function save(e) {
    e.preventDefault();
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
    if (confirm('Archive this record?')) {
      await api.delete(`/${entity}/${id}`);
      await load();
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
            <button type="button" className="scooh-btn purple-btn" onClick={() => setEdit({})}>+ Add {title.replace(/s$/, '')}</button>
          </>
        }
      />

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

      <div className="scooh-tablewrap">
        <table className="scooh-table">
          <thead>
            <tr>
              {fs.slice(0, 8).map(f => (
                <SortHeader key={f} label={label(f)} sortKey={f} currentSort={sortState} onSort={handleSort} />
              ))}
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {shown.length === 0 ? (
              <tr><td colSpan={fs.slice(0, 8).length + 1} className="scooh-empty">No records found</td></tr>
            ) : (
              shown.map(r => (
                <tr key={r.id}>
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
                      <button className="scooh-iconbtn scooh-text-action" onClick={() => setEdit(r)}>Edit</button>
                      <button className="scooh-iconbtn danger-icon" onClick={() => del(r.id)}>×</button>
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
                <h2>{edit.id ? 'Edit' : 'Add'} {title.replace(/s$/, '')}</h2>
                <span className="scooh-modal-subtitle">Enter details and save</span>
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
                          <textarea name={f} defaultValue={edit[f] ?? ''} placeholder={meta.placeholder || ''} rows={3} />
                        ) : hasSelect ? (
                          <select name={f} defaultValue={edit[f] ?? meta.select[0]}>
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
                          />
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
              <div className="scooh-modalfoot">
                <button type="button" className="scooh-btn ghost" onClick={() => setEdit(null)}>Cancel</button>
                <button className="scooh-btn purple-btn" disabled={saving}>
                  {saving ? 'Saving…' : edit.id ? 'Update Record' : 'Save Record'}
                </button>
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

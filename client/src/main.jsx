import React, { useEffect, useMemo, useState, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import { BrowserRouter, Navigate, Route, Routes, useLocation, useNavigate } from 'react-router-dom';
import * as XLSX from 'xlsx';
import PptxGenJS from 'pptxgenjs';
import api from './api';
import './styles.css';

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

const label = s => String(s).replaceAll('_', ' ').replace(/\b\w/g, m => m.toUpperCase());
const money = v => '₹' + Number(v || 0).toLocaleString('en-IN', { maximumFractionDigits: 2 });
const formatDate = v => v ? new Date(v).toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' }) : '—';

function unpackSite(s) {
  let f = s.flags;
  try {
    if (typeof f === 'string') f = JSON.parse(f);
  } catch {}
  f = f && typeof f === 'object' && !Array.isArray(f) ? f : {};
  return {
    ...s,
    ppt_images: Array.isArray(f.ppt_images) ? f.ppt_images : [],
    ppt_availability: f.ppt_availability || '',
    ppt_rate: f.ppt_rate || ''
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
  const user = JSON.parse(localStorage.getItem('sc_user') || '{}');
  const [notifOpen, setNotifOpen] = useState(false);
  const [notifications, setNotifications] = useState([]);
  const [sidebarOpen, setSidebarOpen] = useState(false);

  useEffect(() => {
    api.get('/notifications').then(r => setNotifications(r.data)).catch(() => {});
  }, []);

  const unreadCount = notifications.filter(n => !n.is_read).length;

  async function clearAllNotifs() {
    await api.post('/notifications/read', {});
    setNotifications(notifications.map(n => ({ ...n, is_read: 1 })));
  }

  return (
    <div className="scooh-wp-wrap">
      <div className="scooh-app scooh-has-topbar">
        {/* Top Header */}
        <header className="scooh-topbar">
          <div className="scooh-topbar-left">
            <button className="scooh-mobile-nav-toggle" onClick={() => setSidebarOpen(!sidebarOpen)} aria-label="Toggle Menu">
              <span></span><span></span><span></span>
            </button>
            <div className="scooh-topbar-context">
              <span className="scooh-topbar-eyebrow">MEDIA BUZZ • WORKSPACE</span>
              <strong>OOH Operations Management</strong>
            </div>
          </div>

          <div className="scooh-topbar-user">
            <div className="scooh-topbar-notifications" style={{ position: 'relative' }}>
              <button
                type="button"
                className="scooh-notification-bell scooh-iconbtn"
                onClick={() => setNotifOpen(!notifOpen)}
                aria-label="Notifications"
                style={{ position: 'relative' }}
              >
                <span style={{ fontSize: '15px' }}>🔔</span>
                {unreadCount > 0 && (
                  <span className="scooh-notification-count" style={{ position: 'absolute', top: '-4px', right: '-4px' }}>
                    {unreadCount}
                  </span>
                )}
              </button>

              {notifOpen && (
                <div className="scooh-top-notification-panel" style={{ display: 'block', position: 'absolute', right: 0, top: '44px', zIndex: 1000, width: '320px', background: '#111720', border: '1px solid #27313d', borderRadius: '12px', padding: '14px', boxShadow: '0 12px 30px rgba(0,0,0,0.5)' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
                    <strong style={{ fontSize: '12px', color: '#fff' }}>Notifications</strong>
                    <button type="button" className="scooh-btn ghost" style={{ fontSize: '10px', padding: '4px 8px', minHeight: 'auto' }} onClick={clearAllNotifs}>
                      Mark all read
                    </button>
                  </div>
                  <div style={{ maxHeight: '220px', overflowY: 'auto' }}>
                    {notifications.length === 0 ? (
                      <p style={{ color: '#74808d', fontSize: '11px', margin: 0, padding: '10px 0' }}>No new notifications</p>
                    ) : (
                      notifications.map(n => (
                        <div key={n.id} style={{ padding: '8px 0', borderBottom: '1px solid #1a222c', fontSize: '11px', color: n.is_read ? '#74808d' : '#f0f3f6' }}>
                          <strong>{n.title}</strong>
                          <p style={{ margin: '2px 0 0', color: '#8d98a6' }}>{n.message}</p>
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
                    className={loc.pathname.includes(k) ? 'active' : ''}
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

        {sidebarOpen && <div className="scooh-mobile-scrim" style={{ display: 'block', opacity: 1, visibility: 'visible' }} onClick={() => setSidebarOpen(false)} />}

        {/* Main Content Area */}
        <main className="scooh-main">
          <div className="scooh-view">
            <Routes>
              <Route path="/dashboard" element={<Dashboard />} />
              <Route path="/sites" element={<SitesView />} />
              <Route path="/campaigns" element={<CampaignsView />} />
              <Route path="/occupancy" element={<OccupancyView />} />
              <Route path="/proposals" element={<ProposalsView />} />
              <Route path="/ppt" element={<PptView />} />
              <Route path="/electricity" element={<Crud entity="electricity" title="Electricity" />} />
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
        <PageHead title="Dashboard" desc="Media Buzz OOH operations overview" />
        <section className="scooh-panel">
          <p style={{ color: '#f06a6a', fontWeight: 600 }}>{err}</p>
          <button className="scooh-btn primary" onClick={load}>Retry</button>
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

  return (
    <>
      <PageHead
        title="Dashboard"
        desc="Media Buzz OOH operations overview, live occupancy, alerts, and revenue tracking."
        actions={
          <>
            <button className="scooh-btn primary" onClick={() => nav('/sites')}>+ Sites</button>
            <button className="scooh-btn" onClick={() => nav('/campaigns')}>+ Campaign</button>
            <button className="scooh-btn" onClick={() => nav('/ppt')}>Generate PPT</button>
          </>
        }
      />

      {/* KPI Cards Grid */}
      <div className="scooh-kpirow">
        <div className="scooh-kpi good">
          <div className="n">{k.total_sites || 0}</div>
          <div className="l">Total Sites</div>
          <div className="scooh-kpi-note" style={{ color: '#48c79a' }}>{k.available_sites || 0} Available</div>
        </div>
        <div className="scooh-kpi">
          <div className="n">{k.active_campaigns || 0}</div>
          <div className="l">Active Campaigns</div>
          <div className="scooh-kpi-note">Live on displays</div>
        </div>
        <div className="scooh-kpi good">
          <div className="n">{money(k.campaign_revenue)}</div>
          <div className="l">Total Revenue</div>
          <div className="scooh-kpi-note" style={{ color: '#48c79a' }}>Margin: {money(k.gross_margin)}</div>
        </div>
        <div className={`scooh-kpi ${Number(k.electricity_overdue) > 0 ? 'danger' : 'good'}`}>
          <div className="n">{money(k.unpaid_electricity)}</div>
          <div className="l">Electricity Due</div>
          <div className="scooh-kpi-note" style={{ color: Number(k.electricity_overdue) > 0 ? '#f06a6a' : '#929daa' }}>
            {k.electricity_overdue || 0} Overdue Bills
          </div>
        </div>
      </div>

      {/* Action Alerts Section */}
      <section className="scooh-panel">
        <h3>Action Alerts & Campaign Deadlines</h3>
        <div className="scooh-tablewrap">
          <table className="scooh-table">
            <thead>
              <tr>
                <th>Site ID</th>
                <th>Client</th>
                <th>Campaign Name</th>
                <th>End Date</th>
                <th>Invoice Status</th>
                <th>Hard Copy</th>
              </tr>
            </thead>
            <tbody>
              {!d.alerts || d.alerts.length === 0 ? (
                <tr>
                  <td colSpan="6" className="scooh-empty">No pending action alerts</td>
                </tr>
              ) : (
                d.alerts.map(x => (
                  <tr key={x.id}>
                    <td><span className="scooh-plate">{x.site_code}</span></td>
                    <td><b>{x.client || '—'}</b></td>
                    <td>{x.campaign_name || '—'}</td>
                    <td>{formatDate(x.end_date)}</td>
                    <td>
                      <span className={`scooh-pill ${x.invoice_status === 'Paid' ? 'active' : x.invoice_status === 'Sent' ? 'watch' : 'vacant'}`}>
                        {x.invoice_status || 'Pending'}
                      </span>
                    </td>
                    <td>
                      <span className={`scooh-pill ${x.hard_copy_status === 'Delivered' ? 'active' : x.hard_copy_status === 'Dispatched' ? 'watch' : 'vacant'}`}>
                        {x.hard_copy_status || 'Pending'}
                      </span>
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

function SitesView() {
  const [sites, setSites] = useState([]);
  const [search, setSearch] = useState('');
  const [cityFilter, setCityFilter] = useState('');
  const [mediaFilter, setMediaFilter] = useState('');
  const [edit, setEdit] = useState(null);
  const [imageModalSite, setImageModalSite] = useState(null);

  async function load() {
    const { data } = await api.get('/sites');
    setSites(data.map(unpackSite));
  }

  useEffect(() => { load(); }, []);

  const cities = useMemo(() => Array.from(new Set(sites.map(s => s.city).filter(Boolean))), [sites]);
  const mediaTypes = useMemo(() => Array.from(new Set(sites.map(s => s.media_type).filter(Boolean))), [sites]);

  const filtered = sites.filter(s => {
    const text = JSON.stringify(s).toLowerCase();
    const matchesQuery = text.includes(search.toLowerCase());
    const matchesCity = !cityFilter || s.city === cityFilter;
    const matchesMedia = !mediaFilter || s.media_type === mediaFilter;
    return matchesQuery && matchesCity && matchesMedia;
  });

  async function save(e) {
    e.preventDefault();
    const data = Object.fromEntries(new FormData(e.currentTarget));
    if (edit?.id) await api.put(`/sites/${edit.id}`, data);
    else await api.post('/sites', data);
    setEdit(null);
    await load();
  }

  async function del(id) {
    if (confirm('Archive this site record?')) {
      await api.delete(`/sites/${id}`);
      load();
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
            <button className="scooh-btn primary" onClick={() => setEdit({})}>+ Add Site</button>
          </>
        }
      />

      <div className="scooh-toolbar">
        <input
          className="scooh-search"
          placeholder="Search site ID, location, landmark…"
          value={search}
          onChange={e => setSearch(e.target.value)}
        />
        <div className="scooh-filters">
          <select value={cityFilter} onChange={e => setCityFilter(e.target.value)}>
            <option value="">All Cities ({cities.length})</option>
            {cities.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
          <select value={mediaFilter} onChange={e => setMediaFilter(e.target.value)}>
            <option value="">All Media Types</option>
            {mediaTypes.map(m => <option key={m} value={m}>{m}</option>)}
          </select>
        </div>
      </div>

      <div className="scooh-tablewrap">
        <table className="scooh-table">
          <thead>
            <tr>
              <th>Site ID</th>
              <th>Area / Landmark</th>
              <th>City</th>
              <th>Type</th>
              <th>Size</th>
              <th>Lighting</th>
              <th>Availability</th>
              <th>Monthly Rate</th>
              <th>GPS / Maps</th>
              <th>Images</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 ? (
              <tr><td colSpan="11" className="scooh-empty">No matching sites found</td></tr>
            ) : (
              filtered.map(s => (
                <tr key={s.id}>
                  <td><span className="scooh-plate">{s.site_code}</span></td>
                  <td>
                    <b>{s.area || '—'}</b>
                    {s.address && <div style={{ fontSize: '10px', color: '#74808d' }}>{s.address}</div>}
                  </td>
                  <td>{s.city}</td>
                  <td><span className="scooh-badgechip">{s.media_type}</span></td>
                  <td>{s.size || (s.width && s.height ? `${s.width}x${s.height}` : '—')}</td>
                  <td><span className="scooh-badgechip">{s.lighting}</span></td>
                  <td>
                    <span className={`scooh-pill ${s.availability === 'Available' ? 'active' : 'vacant'}`}>
                      {s.availability}
                    </span>
                  </td>
                  <td><b>{money(s.monthly_rate)}</b></td>
                  <td>
                    {s.gps ? (
                      <a className="scooh-map-link" href={`https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(s.gps)}`} target="_blank" rel="noreferrer">
                        {s.gps}
                      </a>
                    ) : '—'}
                  </td>
                  <td>
                    <div style={{ display: 'flex', gap: '4px', alignItems: 'center' }}>
                      {(s.ppt_images || []).slice(0, 2).map((img, i) => (
                        <img key={i} src={img} alt="Thumb" style={{ width: '28px', height: '28px', objectFit: 'cover', borderRadius: '4px', border: '1px solid #34404e' }} />
                      ))}
                      <button className="scooh-btn ghost" style={{ minHeight: '26px', padding: '2px 6px', fontSize: '10px' }} onClick={() => setImageModalSite(s)}>
                        +{(s.ppt_images || []).length}
                      </button>
                    </div>
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
                    <input name="site_code" defaultValue={edit.site_code ?? ''} required />
                  </div>
                  <div className="scooh-field">
                    <label>City</label>
                    <input name="city" defaultValue={edit.city ?? 'Ahmedabad'} required />
                  </div>
                  <div className="scooh-field">
                    <label>Area / Landmark</label>
                    <input name="area" defaultValue={edit.area ?? ''} required />
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
                    <input name="size" defaultValue={edit.size ?? ''} />
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
                    <input name="monthly_rate" type="number" defaultValue={edit.monthly_rate ?? 0} />
                  </div>
                  <div className="scooh-field">
                    <label>Availability</label>
                    <select name="availability" defaultValue={edit.availability ?? 'Available'}>
                      <option value="Available">Available</option>
                      <option value="Occupied">Occupied</option>
                      <option value="Maintenance">Maintenance</option>
                    </select>
                  </div>
                  <div className="scooh-field" style={{ gridColumn: '1 / -1' }}>
                    <label>Full Address</label>
                    <textarea name="address" defaultValue={edit.address ?? ''} />
                  </div>
                  <div className="scooh-field">
                    <label>GPS Coordinates</label>
                    <input name="gps" defaultValue={edit.gps ?? ''} placeholder="23.019187, 72.530184" />
                  </div>
                  <div className="scooh-field">
                    <label>Google Maps URL</label>
                    <input name="maps_url" defaultValue={edit.maps_url ?? ''} />
                  </div>
                </div>
              </div>
              <div className="scooh-modalfoot">
                <button type="button" className="scooh-btn ghost" onClick={() => setEdit(null)}>Cancel</button>
                <button className="scooh-btn primary">Save Site</button>
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
  const [sites, setSites] = useState([]);
  const [pages, setPages] = useState({});
  const [sel, setSel] = useState({});
  const [query, setQuery] = useState('');
  const [generating, setGenerating] = useState(false);

  async function load() {
    const [sRes, pRes] = await Promise.all([
      api.get('/sites'),
      api.get('/ppt-pages')
    ]);
    setSites(sRes.data.map(unpackSite));
    setPages(pRes.data);
  }

  useEffect(() => { load(); }, []);

  const filtered = sites.filter(s => JSON.stringify(s).toLowerCase().includes(query.toLowerCase()));

  async function uploadPage(k, file) {
    const fd = new FormData();
    fd.append('file', file);
    await api.post('/ppt-pages/' + k, fd, { headers: { 'Content-Type': 'multipart/form-data' } });
    load();
  }

  async function addImages(id, files) {
    const fd = new FormData();
    [...files].forEach(f => fd.append('files', f));
    await api.post(`/sites/${id}/images`, fd, { headers: { 'Content-Type': 'multipart/form-data' } });
    load();
  }

  async function generate() {
    const chosen = sites
      .filter(s => sel[s.id]?.checked)
      .map(s => ({
        ...s,
        _availability: sel[s.id]?.availability ?? s.ppt_availability,
        _rate: sel[s.id]?.rate ?? s.ppt_rate,
        _showRate: !!sel[s.id]?.showRate,
        _showCoords: !!sel[s.id]?.showCoords
      }));

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
        title="Automated PPT Presentation Generator"
        desc="Generate customer-ready PPTs using Media Buzz format: Fixed Page 1 & 2 -> Selected Site Image Slides -> Fixed Page 3."
        actions={
          <button className="scooh-btn primary" onClick={generate} disabled={generating}>
            {generating ? 'Creating PPT…' : 'Generate Presentation (.pptx)'}
          </button>
        }
      />

      {/* Fixed Pages Cards */}
      <section className="scooh-panel">
        <h3>Fixed Presentation Templates</h3>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '14px' }}>
          {[
            ['first', 'Page 1 — Cover Slide'],
            ['second_last', 'Page 2 — Overview Slide'],
            ['last', 'Page 3 — Thank You Slide']
          ].map(([k, title]) => (
            <label key={k} style={{ display: 'block', padding: '12px', border: '1px solid #27313d', borderRadius: '12px', background: '#111720', cursor: 'pointer' }}>
              <div style={{ height: '130px', background: '#080c11', borderRadius: '8px', overflow: 'hidden', display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: '8px' }}>
                {pages[k] ? (
                  <img src={pages[k]} alt={title} style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
                ) : (
                  <span style={{ color: '#74808d', fontSize: '11px' }}>Click to upload {title}</span>
                )}
              </div>
              <strong style={{ fontSize: '11px', color: '#f4f6f8' }}>{title}</strong>
              <input type="file" accept="image/*" hidden onChange={e => e.target.files[0] && uploadPage(k, e.target.files[0])} />
            </label>
          ))}
        </div>
      </section>

      {/* Sites Toolbar */}
      <div className="scooh-toolbar">
        <input
          className="scooh-search"
          placeholder="Filter sites for presentation…"
          value={query}
          onChange={e => setQuery(e.target.value)}
        />
        <div className="scooh-filters">
          <button
            type="button"
            className="scooh-btn"
            onClick={() => setSel(Object.fromEntries(filtered.map(s => [s.id, { ...(sel[s.id] || {}), checked: true }])))}
          >
            Select All Visible ({filtered.length})
          </button>
          <button
            type="button"
            className="scooh-btn ghost"
            onClick={() => setSel({})}
          >
            Deselect All
          </button>
        </div>
      </div>

      {/* Selected Site Cards */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: '14px' }}>
        {filtered.map(s => {
          const v = sel[s.id] || {};
          const isChecked = !!v.checked;

          return (
            <article
              key={s.id}
              style={{
                padding: '16px',
                border: isChecked ? '1px solid #f2c94c' : '1px solid #27313d',
                borderRadius: '14px',
                background: isChecked ? 'rgba(242, 201, 76, 0.05)' : '#131922'
              }}
            >
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={isChecked}
                    onChange={e => setSel({ ...sel, [s.id]: { ...v, checked: e.target.checked } })}
                  />
                  <span className="scooh-plate">{s.site_code}</span>
                </label>
                <span className="scooh-badgechip">{s.media_type}</span>
              </div>

              <p style={{ margin: '4px 0 10px', fontSize: '12px', fontWeight: 600, color: '#f0f3f6' }}>
                {s.area || s.address || s.city}
              </p>

              {/* Photos preview */}
              <div style={{ display: 'flex', gap: '6px', overflowX: 'auto', marginBottom: '10px' }}>
                {(s.ppt_images || []).map((u, idx) => (
                  <img key={idx} src={u} alt="Site" style={{ width: '56px', height: '40px', objectFit: 'cover', borderRadius: '6px', border: '1px solid #34404e' }} />
                ))}
              </div>

              <label className="scooh-btn ghost" style={{ width: '100%', marginBottom: '10px', fontSize: '11px', cursor: 'pointer' }}>
                + Upload Images
                <input type="file" multiple accept="image/*" hidden onChange={e => addImages(s.id, e.target.files)} />
              </label>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                <label style={{ fontSize: '10px', color: '#8d98a6' }}>
                  Availability
                  <input
                    value={v.availability ?? s.ppt_availability}
                    onChange={e => setSel({ ...sel, [s.id]: { ...v, availability: e.target.value } })}
                    style={{ marginTop: '2px' }}
                  />
                </label>
                <label style={{ fontSize: '10px', color: '#8d98a6' }}>
                  Rate per Month (₹)
                  <input
                    value={v.rate ?? s.ppt_rate}
                    onChange={e => setSel({ ...sel, [s.id]: { ...v, rate: e.target.value } })}
                    style={{ marginTop: '2px' }}
                  />
                </label>
                <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '10.5px', color: '#cdd5df', marginTop: '4px' }}>
                  <input
                    type="checkbox"
                    checked={!!v.showRate}
                    onChange={e => setSel({ ...sel, [s.id]: { ...v, showRate: e.target.checked } })}
                  />
                  Show Rate in Slide
                </label>
                <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '10.5px', color: '#cdd5df' }}>
                  <input
                    type="checkbox"
                    checked={!!v.showCoords}
                    onChange={e => setSel({ ...sel, [s.id]: { ...v, showCoords: e.target.checked } })}
                  />
                  Show Coordinates
                </label>
              </div>
            </article>
          );
        })}
      </div>
    </>
  );
}

function OccupancyView() {
  const [rows, setRows] = useState([]);

  useEffect(() => {
    Promise.all([api.get('/sites'), api.get('/campaigns')]).then(([s, c]) => {
      const now = new Date(), from = new Date(now);
      from.setDate(now.getDate() - 365);
      const calculated = s.data.map(site => {
        let days = 0;
        c.data.filter(x => String(x.site_id) === String(site.id)).forEach(x => {
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
      }).sort((a, b) => b.pct - a.pct);
      setRows(calculated);
    });
  }, []);

  return (
    <>
      <PageHead title="Occupancy & Utilization" desc="Rolling 365-day site occupancy calculations and performance breakdown." />
      <section className="scooh-panel">
        <div className="scooh-tablewrap">
          <table className="scooh-table">
            <thead>
              <tr>
                <th>Site ID</th>
                <th>City</th>
                <th>Area</th>
                <th>Occupied Days (365d)</th>
                <th>Occupancy %</th>
              </tr>
            </thead>
            <tbody>
              {rows.map(r => (
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
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </>
  );
}

function CampaignsView() {
  return <Crud entity="campaigns" title="Campaign Tracker" />;
}

function ProposalsView() {
  const [sites, setSites] = useState([]);
  const [clients, setClients] = useState([]);
  const [selectedClient, setSelectedClient] = useState('');
  const [campaignName, setCampaignName] = useState('');
  const [proposalDate, setProposalDate] = useState(new Date().toISOString().slice(0, 10));
  const [validityDays, setValidityDays] = useState(7);
  const [discountPercent, setDiscountPercent] = useState(0);
  const [taxPercent, setTaxPercent] = useState(18);
  const [selectedSites, setSelectedSites] = useState({});

  useEffect(() => {
    Promise.all([api.get('/sites'), api.get('/clients')]).then(([s, c]) => {
      setSites(s.data.map(unpackSite));
      setClients(c.data);
    });
  }, []);

  const chosenList = sites.filter(s => selectedSites[s.id]?.checked);
  const subtotal = chosenList.reduce((acc, s) => acc + Number(selectedSites[s.id]?.rate || s.monthly_rate || 0), 0);
  const discountAmount = (subtotal * Number(discountPercent || 0)) / 100;
  const taxableAmount = subtotal - discountAmount;
  const taxAmount = (taxableAmount * Number(taxPercent || 18)) / 100;
  const grandTotal = taxableAmount + taxAmount;

  const clientObj = clients.find(c => String(c.id) === String(selectedClient));

  return (
    <>
      <PageHead
        title="Proposal Builder"
        desc="Generate commercial proposals with rates, taxes, client letterhead, and instant print preview."
        actions={
          <button className="scooh-btn primary" onClick={() => window.print()}>Print / Save PDF</button>
        }
      />

      <div className="scooh-proposal-layout">
        {/* Left Controls */}
        <div className="scooh-proposal-controls">
          <div className="scooh-proposal-step">
            <h3>1. Proposal & Client Details</h3>
            <div className="scooh-field">
              <label>Client</label>
              <select value={selectedClient} onChange={e => setSelectedClient(e.target.value)}>
                <option value="">Select a Client</option>
                {clients.map(c => <option key={c.id} value={c.id}>{c.client_name} ({c.company})</option>)}
              </select>
            </div>
            <div className="scooh-field">
              <label>Campaign Name</label>
              <input value={campaignName} onChange={e => setCampaignName(e.target.value)} placeholder="e.g. Diwali Outdoor Campaign" />
            </div>
            <div className="scooh-grid2">
              <div className="scooh-field">
                <label>Proposal Date</label>
                <input type="date" value={proposalDate} onChange={e => setProposalDate(e.target.value)} />
              </div>
              <div className="scooh-field">
                <label>Validity (Days)</label>
                <input type="number" value={validityDays} onChange={e => setValidityDays(e.target.value)} />
              </div>
            </div>
          </div>

          <div className="scooh-proposal-step">
            <h3>2. Select Sites & Commercials</h3>
            <div className="scooh-picklist">
              {sites.map(s => {
                const isChecked = !!selectedSites[s.id]?.checked;
                return (
                  <div key={s.id} className="scooh-pickrow">
                    <input
                      type="checkbox"
                      checked={isChecked}
                      onChange={e => setSelectedSites({
                        ...selectedSites,
                        [s.id]: {
                          checked: e.target.checked,
                          rate: selectedSites[s.id]?.rate ?? s.monthly_rate
                        }
                      })}
                    />
                    <span className="scooh-plate">{s.site_code}</span>
                    <div className="scooh-pickmeta">
                      <strong>{s.area || s.city}</strong>
                      <span>{s.media_type} · {s.size}</span>
                    </div>
                    {isChecked && (
                      <input
                        type="number"
                        value={selectedSites[s.id]?.rate ?? s.monthly_rate}
                        onChange={e => setSelectedSites({
                          ...selectedSites,
                          [s.id]: { ...selectedSites[s.id], rate: e.target.value }
                        })}
                        placeholder="Rate ₹"
                      />
                    )}
                  </div>
                );
              })}
            </div>
          </div>

          <div className="scooh-proposal-step">
            <h3>3. Taxes & Adjustments</h3>
            <div className="scooh-grid2">
              <div className="scooh-field">
                <label>Discount %</label>
                <input type="number" value={discountPercent} onChange={e => setDiscountPercent(e.target.value)} />
              </div>
              <div className="scooh-field">
                <label>GST Tax %</label>
                <input type="number" value={taxPercent} onChange={e => setTaxPercent(e.target.value)} />
              </div>
            </div>
          </div>
        </div>

        {/* Right Live Document Preview */}
        <div className="scooh-proposal-document">
          <div className="scooh-doc">
            <div className="scooh-proposal-letterhead">
              <img className="scooh-proposal-logo" src="/assets/media-buzz-logo.png" alt="Media Buzz" />
              <div className="scooh-doc-date">
                <strong>PROPOSAL CODE: MB-PROP-{Math.floor(Date.now() / 100000)}</strong>
                <div>Date: {formatDate(proposalDate)}</div>
                <div>Valid for: {validityDays} Days</div>
              </div>
            </div>

            <div className="scooh-doc-rule" />

            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '18px' }}>
              <div>
                <span style={{ fontSize: '10px', color: '#667', textTransform: 'uppercase', fontWeight: 800 }}>Prepared For:</span>
                <h3 style={{ margin: '4px 0 2px', fontSize: '15px' }}>{clientObj?.client_name || 'Client Name'}</h3>
                <div style={{ fontSize: '11px', color: '#445' }}>{clientObj?.company}</div>
                <div style={{ fontSize: '10px', color: '#667' }}>GSTIN: {clientObj?.gst_number || '—'}</div>
              </div>
              <div style={{ textAlign: 'right' }}>
                <span style={{ fontSize: '10px', color: '#667', textTransform: 'uppercase', fontWeight: 800 }}>Campaign:</span>
                <h3 style={{ margin: '4px 0 2px', fontSize: '15px' }}>{campaignName || 'Outdoor Media Campaign'}</h3>
              </div>
            </div>

            <table>
              <thead>
                <tr>
                  <th>Site ID</th>
                  <th>Location / Landmark</th>
                  <th>Media Type</th>
                  <th>Size</th>
                  <th>Rate / Month (₹)</th>
                </tr>
              </thead>
              <tbody>
                {chosenList.length === 0 ? (
                  <tr><td colSpan="5" className="scooh-doc-empty">Select sites from the left panel to build the proposal.</td></tr>
                ) : (
                  chosenList.map(s => (
                    <tr key={s.id}>
                      <td><strong>{s.site_code}</strong></td>
                      <td>{s.area} - {s.address}</td>
                      <td>{s.media_type}</td>
                      <td>{s.size}</td>
                      <td>{money(selectedSites[s.id]?.rate || s.monthly_rate)}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>

            <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: '18px' }}>
              <div style={{ width: '260px', fontSize: '11.5px', lineHeight: '1.8' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                  <span>Subtotal:</span>
                  <strong>{money(subtotal)}</strong>
                </div>
                {Number(discountPercent) > 0 && (
                  <div style={{ display: 'flex', justifyContent: 'space-between', color: '#c00' }}>
                    <span>Discount ({discountPercent}%):</span>
                    <strong>-{money(discountAmount)}</strong>
                  </div>
                )}
                <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                  <span>GST ({taxPercent}%):</span>
                  <strong>{money(taxAmount)}</strong>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between', borderTop: '2px solid #111', paddingTop: '6px', fontSize: '14px' }}>
                  <strong>Grand Total:</strong>
                  <strong>{money(grandTotal)}</strong>
                </div>
              </div>
            </div>

            <div style={{ marginTop: '28px', borderTop: '1px solid #ddd', paddingTop: '14px', fontSize: '10px', color: '#667' }}>
              <strong>Terms & Conditions:</strong>
              <p style={{ margin: '4px 0' }}>
                1. Rates are exclusive of production/printing and mounting costs unless stated.
                2. 50% advance on confirmation, balance before mounting.
                3. Site availability is subject to confirmation in writing.
              </p>
            </div>
          </div>
        </div>
      </div>
    </>
  );
}

function DataToolsView() {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);

  async function imp() {
    if (!file) return alert('Please choose an Excel file (.xlsx) first.');
    setLoading(true);
    try {
      const f = new FormData();
      f.append('file', file);
      const r = await api.post('/import/xlsx', f, { headers: { 'Content-Type': 'multipart/form-data' } });
      alert(`Success: ${r.data.rows} sites imported successfully.`);
    } catch (e) {
      alert('Import failed: ' + (e.response?.data?.message || e.message));
    } finally {
      setLoading(false);
    }
  }

  async function exportExcel() {
    const { data } = await api.get('/sites');
    const rows = data.map(unpackSite).map(x => ({
      'Site ID': x.site_code,
      'City': x.city,
      'Area / Landmark': x.area,
      'Full Address': x.address,
      'Size': x.size,
      'Width (ft)': x.width,
      'Height (ft)': x.height,
      'Media Type': x.media_type,
      'Lighting': x.lighting,
      'Availability': x.availability,
      'Selling Amount (₹)': x.monthly_rate,
      'PPT Availability': x.ppt_availability,
      'PPT Rate Per Month (₹)': x.ppt_rate,
      'Latitude': x.latitude,
      'Longitude': x.longitude,
      'Google Maps': x.maps_url,
      'Notes': x.notes
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Sites');
    XLSX.writeFile(wb, 'MediaBuzz-Sites-Inventory.xlsx');
  }

  async function exportJson() {
    const { data } = await api.get('/export/json');
    const b = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const u = URL.createObjectURL(b);
    const a = document.createElement('a');
    a.href = u;
    a.download = 'mediabuzz-database-backup.json';
    a.click();
    URL.revokeObjectURL(u);
  }

  return (
    <>
      <PageHead title="Import & Export Tools" desc="Excel inventory migration, WordPress data import, and database backup downloads." />
      <div className="scooh-data-grid">
        <section className="scooh-data-panel">
          <h3>Import Excel Inventory</h3>
          <p className="scooh-data-copy">
            Import Excel (.xlsx) files. Automatically maps Site ID, AREA, LOCATION, MEDIA, LIGHT, Width, Height, Selling Amount, PPT Availability, Latitude, and Longitude.
          </p>
          <div className="scooh-file-control">
            <input type="file" accept=".xlsx,.xls" onChange={e => setFile(e.target.files[0])} />
          </div>
          <div className="scooh-data-actions">
            <button className="scooh-btn primary" onClick={imp} disabled={loading}>
              {loading ? 'Importing…' : 'Import Excel'}
            </button>
          </div>
        </section>

        <section className="scooh-data-panel">
          <h3>Export & Backup</h3>
          <p className="scooh-data-copy">
            Download your active sites database formatted for Excel, or export a complete JSON snapshot of all Media Buzz tables.
          </p>
          <div className="scooh-data-actions">
            <button className="scooh-btn" onClick={exportExcel}>Export Sites to Excel</button>
            <button className="scooh-btn ghost" onClick={exportJson}>Download JSON Backup</button>
          </div>
        </section>
      </div>
    </>
  );
}

function ReportsView() {
  return <Dashboard />;
}

function SettingsView() {
  const [s, setS] = useState({});
  const [saved, setSaved] = useState(false);

  useEffect(() => {
    api.get('/settings').then(r => setS(r.data));
  }, []);

  async function save(e) {
    e.preventDefault();
    const data = Object.fromEntries(new FormData(e.currentTarget));
    await api.put('/settings', data);
    setSaved(true);
    setTimeout(() => setSaved(false), 3000);
  }

  return (
    <>
      <PageHead title="Settings & Parameters" desc="Configure company details, default GST tax rates, reminder thresholds, and terms." />
      <form className="scooh-panel" onSubmit={save}>
        {saved && <div className="scooh-banner" style={{ display: 'block' }}>Settings saved successfully!</div>}
        <div className="scooh-grid2">
          {[
            ['company_name', 'Company Name'],
            ['currency', 'Currency (e.g. INR)'],
            ['default_city', 'Default City'],
            ['validation_interval', 'Validation Interval (Days)'],
            ['mounting_grace_days', 'Mounting Grace Days'],
            ['electricity_due_soon_days', 'Electricity Due Warning (Days)'],
            ['campaign_ending_warning_days', 'Campaign End Warning (Days)'],
            ['notification_emails', 'Notification Recipient Emails'],
            ['proposal_validity', 'Proposal Default Validity (Days)'],
            ['proposal_tax', 'Default Proposal Tax / GST %']
          ].map(([k, lbl]) => (
            <div className="scooh-field" key={k}>
              <label>{lbl}</label>
              <input name={k} defaultValue={s[k] ?? ''} />
            </div>
          ))}
          <div className="scooh-field" style={{ gridColumn: '1 / -1' }}>
            <label>Default Proposal Terms & Conditions</label>
            <textarea name="proposal_terms" defaultValue={s.proposal_terms ?? ''} />
          </div>
        </div>
        <button className="scooh-btn primary" style={{ marginTop: '16px' }}>Save Settings</button>
      </form>
    </>
  );
}

function Crud({ entity, title }) {
  const [rows, setRows] = useState([]);
  const [edit, setEdit] = useState(null);
  const [search, setSearch] = useState('');
  const fs = fields[entity] || [];

  async function load() {
    const { data } = await api.get('/' + entity);
    setRows(data);
  }

  useEffect(() => { load(); }, [entity]);

  const shown = rows.filter(r => JSON.stringify(r).toLowerCase().includes(search.toLowerCase()));

  async function save(e) {
    e.preventDefault();
    const data = Object.fromEntries(new FormData(e.currentTarget));
    if (edit?.id) await api.put(`/${entity}/${edit.id}`, data);
    else await api.post('/' + entity, data);
    setEdit(null);
    await load();
  }

  async function del(id) {
    if (confirm('Archive this record?')) {
      await api.delete(`/${entity}/${id}`);
      load();
    }
  }

  return (
    <>
      <PageHead
        title={title}
        desc={`Managing ${shown.length} total ${title.toLowerCase()} records.`}
        actions={
          <button className="scooh-btn primary" onClick={() => setEdit({})}>+ Add {title.replace(/s$/, '')}</button>
        }
      />

      <div className="scooh-toolbar">
        <input
          className="scooh-search"
          placeholder={`Search ${title.toLowerCase()}…`}
          value={search}
          onChange={e => setSearch(e.target.value)}
        />
      </div>

      <div className="scooh-tablewrap">
        <table className="scooh-table">
          <thead>
            <tr>
              {fs.slice(0, 8).map(f => <th key={f}>{label(f)}</th>)}
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
                        <span className={`scooh-pill ${r[f] === 'Active' || r[f] === 'Paid' || r[f] === 'Available' ? 'active' : r[f] === 'Pending' ? 'watch' : 'vacant'}`}>
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
                  {fs.map(f => (
                    <div className="scooh-field" key={f} style={f === 'notes' || f === 'address' || f === 'billing_address' || f === 'terms' ? { gridColumn: '1 / -1' } : {}}>
                      <label>{label(f)}</label>
                      {f === 'notes' || f === 'address' || f === 'billing_address' || f === 'terms' ? (
                        <textarea name={f} defaultValue={edit[f] ?? ''} />
                      ) : (
                        <input
                          name={f}
                          defaultValue={edit[f] ?? ''}
                          type={/date/.test(f) ? 'date' : /cost|rate|amount|revenue|width|height|tax|discount|days|site_id|client_id/.test(f) ? 'number' : 'text'}
                          step="any"
                        />
                      )}
                    </div>
                  ))}
                </div>
              </div>
              <div className="scooh-modalfoot">
                <button type="button" className="scooh-btn ghost" onClick={() => setEdit(null)}>Cancel</button>
                <button className="scooh-btn primary">Save</button>
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
    api.get('/' + endpoint).then(r => setRows(r.data));
  }, [endpoint]);

  return (
    <>
      <PageHead title={title} desc={`Reviewing recent system ${title.toLowerCase()} events.`} />
      <div className="scooh-tablewrap">
        <table className="scooh-table">
          <thead>
            <tr>
              {Object.keys(rows[0] || {}).slice(0, 9).map(k => <th key={k}>{label(k)}</th>)}
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
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
    <Routes>
      <Route path="/login" element={<Login />} />
      <Route path="/*" element={<Guard><Layout /></Guard>} />
    </Routes>
  );
}

createRoot(document.getElementById('root')).render(
  <BrowserRouter>
    <App />
  </BrowserRouter>
);

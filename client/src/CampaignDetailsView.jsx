import React, { useState, useEffect, useMemo } from 'react';
import { useNavigate, useLocation } from 'react-router-dom';
import api from './api';
import ExcelJS from 'exceljs';
import { attachMediaBuzzExcelHeader, attachMediaBuzzTermsAndConditions } from './excelLogo';

export function money(val) {
  const n = Number(val || 0);
  if (!n) return '₹0';
  return '₹' + n.toLocaleString('en-IN');
}

export function formatDate(val) {
  if (!val) return '';
  const d = new Date(val);
  if (isNaN(d.getTime())) return String(val);
  return d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
}

function isVacantClient(clientName) {
  if (!clientName) return true;
  const s = String(clientName).trim().toLowerCase();
  return (
    s === '' ||
    s === 'vacant' ||
    s === 'empty' ||
    s === 'none' ||
    s === 'available' ||
    s === 'self' ||
    s === 'media buzz' ||
    s === 'media buzz vacant' ||
    s === '0% vacant' ||
    s.includes('vacant') ||
    s === '—' ||
    s === '-'
  );
}

// ── Export Filtered Campaign Details to Excel ──────────────────────────────
async function exportCampaignDetailsExcel(rowsData, filters = {}) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Campaign Details', {
    views: [{ state: 'frozen', ySplit: 2 }]
  });

  const cols = [
    { key: 'sr_no', width: 8 },
    { key: 'site_code', width: 14 },
    { key: 'client_display', width: 36 },
    { key: 'location', width: 38 },
    { key: 'size', width: 14 },
    { key: 'type', width: 16 },
    { key: 'start_date', width: 15 },
    { key: 'end_date', width: 15 },
    { key: 'days', width: 10 }
  ];
  worksheet.columns = cols;

  let title = 'MEDIA BUZZ — CAMPAIGN DETAILS REPORT';
  if (filters.clientOrDisplay) {
    title += ` • ${filters.clientOrDisplay.toUpperCase()}`;
  }
  if (filters.startDate || filters.endDate) {
    title += ` [${filters.startDate || 'Any'} to ${filters.endDate || 'Any'}]`;
  }

  // Executive Media Buzz Brand Header Banner with Logo
  attachMediaBuzzExcelHeader(workbook, worksheet, {
    title,
    columns: cols,
    totalColumns: cols.length
  });

  // Header Row (Row 2)
  const headers = [
    '#', 'Site Code', 'Client / Display',
    'Location', 'Size', 'Type', 'Start Date', 'End Date', 'Days'
  ];
  const headerRow = worksheet.getRow(2);
  headerRow.height = 28;
  headers.forEach((h, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = h;
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF000000' } };
    cell.alignment = {
      vertical: 'middle',
      horizontal: [1, 2, 5, 6, 7, 8, 9].includes(i + 1) ? 'center' : 'left',
      wrapText: false
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFB0B0B0' } },
      left: { style: 'thin', color: { argb: 'FFB0B0B0' } },
      bottom: { style: 'medium', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FFB0B0B0' } }
    };
  });

  // Populate data rows
  rowsData.forEach((r, idx) => {
    const clientName = r.client || r.client_name || '';
    const disp = r.display || r.campaign_name || r.brand || '';
    const combinedClientDisplay = clientName && disp && clientName !== disp ? `${clientName} (${disp})` : (clientName || disp || '—');

    const row = worksheet.addRow({
      sr_no: idx + 1,
      site_code: r.site_code || '—',
      client_display: combinedClientDisplay,
      location: r.location || '—',
      size: r.size || (r.width && r.height ? `${r.width}x${r.height} ft` : '—'),
      type: r.type || 'Hoarding',
      start_date: r.start_date || '—',
      end_date: r.end_date || '—',
      days: r.tenureDays || r.days || (r.start_date && r.end_date ? Math.max(1, Math.round((new Date(r.end_date) - new Date(r.start_date)) / 86400000) + 1) : '—')
    });
    row.height = 22;

    const isEven = idx % 2 === 1;
    row.eachCell((cell, colNumber) => {
      cell.font = { name: 'Calibri', size: 10.5 };
      cell.alignment = {
        vertical: 'middle',
        horizontal: [1, 2, 5, 6, 7, 8, 9].includes(colNumber) ? 'center' : 'left'
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
      };
      if (isEven) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF8FAFC' }
        };
      }
    });
  });

  // Enable Auto-Filter on Row 2
  if (rowsData.length > 0) {
    worksheet.autoFilter = `A2:O${rowsData.length + 2}`;
  }

  // Official Media Buzz Terms & Conditions at Bottom Center
  attachMediaBuzzTermsAndConditions(worksheet, {
    totalColumns: cols.length
  });

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `MediaBuzz_Campaign_Details_${new Date().toISOString().slice(0, 10)}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}


// ── Download Sample Template for Campaign Excel Import ───────────────────────
export async function downloadSampleCampaignExcel() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Media Buzz';
  workbook.lastModifiedBy = 'Media Buzz Operations';
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet('Campaign Schedule', {
    views: [{ state: 'normal' }],
    properties: { defaultRowHeight: 20 }
  });

  worksheet.columns = [
    { key: 'colA', width: 14 },
    { key: 'colB', width: 22 },
    { key: 'colC', width: 22 },
    { key: 'colD', width: 44 }
  ];

  const navyDark = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1B2A4A' } };
  const yellowFont = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFF00' } };
  const subheadFont = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF000000' } };
  const dataFont = { name: 'Calibri', size: 10.5, color: { argb: 'FF000000' } };
  const thinBorder = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } }
  };

  // Helper to add a site block matching the exact user format
  function addSiteBlock(siteCode, locationTitle, bookings) {
    // 1. Site Header Row (Dark navy background, yellow bold text)
    const headerRow = worksheet.addRow([siteCode, locationTitle, '', '']);
    headerRow.height = 26;
    headerRow.getCell(1).fill = navyDark;
    headerRow.getCell(1).font = yellowFont;
    headerRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
    headerRow.getCell(1).border = thinBorder;

    headerRow.getCell(2).fill = navyDark;
    headerRow.getCell(2).font = yellowFont;
    headerRow.getCell(2).alignment = { vertical: 'middle', horizontal: 'left' };
    headerRow.getCell(2).border = thinBorder;

    headerRow.getCell(3).fill = navyDark;
    headerRow.getCell(3).border = thinBorder;
    headerRow.getCell(4).fill = navyDark;
    headerRow.getCell(4).border = thinBorder;

    // Merge B, C, D for wide location title
    const headerRowNum = headerRow.number;
    worksheet.mergeCells(`B${headerRowNum}:D${headerRowNum}`);

    // 2. Subheader Row (Up Date | Down Date | Display)
    const subheadRow = worksheet.addRow(['', 'Up Date', 'Down Date', 'Display']);
    subheadRow.height = 22;
    subheadRow.getCell(1).border = thinBorder;

    [2, 3].forEach(col => {
      const cell = subheadRow.getCell(col);
      cell.font = subheadFont;
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = thinBorder;
    });
    const dispCell = subheadRow.getCell(4);
    dispCell.font = subheadFont;
    dispCell.alignment = { vertical: 'middle', horizontal: 'left' };
    dispCell.border = thinBorder;

    // 3. Booking Rows
    bookings.forEach(bk => {
      const bRow = worksheet.addRow(['', bk.upDate, bk.downDate, bk.display]);
      bRow.height = 21;
      bRow.getCell(1).border = thinBorder;

      [2, 3].forEach(col => {
        const cell = bRow.getCell(col);
        cell.font = dataFont;
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = thinBorder;
      });

      const dCell = bRow.getCell(4);
      dCell.font = dataFont;
      dCell.alignment = { vertical: 'middle', horizontal: 'left' };
      dCell.border = thinBorder;
    });

    // Blank separator row
    const emptyRow = worksheet.addRow(['', '', '', '']);
    emptyRow.height = 14;
  }

  // Site 1 (From User Screenshot)
  addSiteBlock(
    'MB-01',
    'Shivranjani Bridge- Nr.D Mart Junction - Left (1) - 30X10 BL',
    [
      { upDate: '28.07.2026', downDate: '31.07.2026', display: 'Swagat Group (Renewal)' },
      { upDate: '01.08.2026', downDate: '30.08.2026', display: 'Swagat Group (Renewal)' },
      { upDate: '01.09.2026', downDate: '30.09.2026', display: 'Swagat Group (Renewal)' }
    ]
  );

  // Site 2 (Example multi-site block)
  addSiteBlock(
    'MB-02',
    'Iskcon Cross Roads - SG Highway - 40X20 BL',
    [
      { upDate: '01.08.2026', downDate: '31.08.2026', display: 'Tata Motors (Tata Punch EV)' },
      { upDate: '01.09.2026', downDate: '30.09.2026', display: 'Tata Motors (Tata Punch EV)' }
    ]
  );

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'MediaBuzz_Campaign_Details_Sample_Template.xlsx';
  a.click();
  URL.revokeObjectURL(url);
}

export default function CampaignDetailsView() {
  const navigate = useNavigate();
  const location = useLocation();
  const user = JSON.parse(localStorage.getItem('sc_user') || '{}');
  const userRole = (user.role || 'admin').toLowerCase();
  const canAdd = ['admin', 'manager', 'editor', 'field_staff'].includes(userRole);
  const canDelete = ['admin', 'manager'].includes(userRole);

  const [campaigns, setCampaigns] = useState([]);
  const [sites, setSites] = useState([]);
  const [loading, setLoading] = useState(true);
  const [importingExcel, setImportingExcel] = useState(false);
  const [importBanner, setImportBanner] = useState('');

  // Filters: Client/Display text, Start Date, End Date, Status
  const [clientOrDisplay, setClientOrDisplay] = useState('');
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [statusFilter, setStatusFilter] = useState('ALL'); // 'ALL' | 'live' | 'upcoming' | 'completed'
  const [siteFilter, setSiteFilter] = useState('');
  const [sortState, setSortState] = useState({ key: 'start_date', dir: 'desc' });


  // ── Import Campaign Excel ────────────────────────────────────────────────
  async function handleImportExcel(e) {
    if (!canAdd) return;
    const file = e.target.files?.[0];
    if (!file) return;
    setImportingExcel(true);
    try {
      const fd = new FormData();
      fd.append('file', file);
      const r = await api.post('/import/campaigns-xlsx', fd, { headers: { 'Content-Type': 'multipart/form-data' } });
      const msg = r.data?.message || `Processed ${r.data?.rows || 0} campaigns (${r.data?.updated || 0} updated, ${r.data?.created || 0} created).`;
      setImportBanner(`✓ ${msg}`);
      alert(`✓ Campaign Excel Import Successful!\n\n${msg}`);
      await loadData();
      window.dispatchEvent(new CustomEvent('mb-campaigns-updated'));
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
      setTimeout(() => setImportBanner(''), 6000);
    } catch (err) {
      console.error('Excel Import failed:', err);
      alert('Campaign Excel Import failed: ' + (err.response?.data?.message || err.message));
    } finally {
      setImportingExcel(false);
      e.target.value = '';
    }
  }

  // Modals: Inspect details and Add/Edit
  const [selectedCampaign, setSelectedCampaign] = useState(null);
  const [editModal, setEditModal] = useState(null);
  const [modalSiteCode, setModalSiteCode] = useState('');
  const [modalClient, setModalClient] = useState('');
  const [modalDisplay, setModalDisplay] = useState('');
  const [modalStartDate, setModalStartDate] = useState('');
  const [modalEndDate, setModalEndDate] = useState('');
  const [modalAdvtFees, setModalAdvtFees] = useState('');
  const [modalProdCost, setModalProdCost] = useState('');
  const [modalTotalAmount, setModalTotalAmount] = useState('');
  const [modalPending, setModalPending] = useState('');
  const [modalNotes, setModalNotes] = useState('');
  const [saving, setSaving] = useState(false);

  // Load Data
  async function loadData() {
    setLoading(true);
    try {
      const [campRes, sitesRes] = await Promise.all([
        api.get('/campaigns').catch(() => ({ data: [] })),
        api.get('/sites').catch(() => ({ data: [] }))
      ]);
      setCampaigns(Array.isArray(campRes.data) ? campRes.data : []);
      setSites(Array.isArray(sitesRes.data) ? sitesRes.data : []);
    } catch (err) {
      console.error('Failed to load campaign details:', err);
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    loadData();
    const interval = setInterval(loadData, 45000);
    return () => clearInterval(interval);
  }, []);

  // Sync with URL query params
  useEffect(() => {
    const params = new URLSearchParams(location.search);
    const qParam = params.get('q') || params.get('client') || params.get('search') || '';
    const startParam = params.get('start') || '';
    const endParam = params.get('end') || '';
    const siteParam = params.get('site') || '';

    if (qParam) setClientOrDisplay(qParam);
    if (startParam) setStartDate(startParam);
    if (endParam) setEndDate(endParam);
    if (siteParam) setSiteFilter(siteParam);
  }, [location.search]);

  // Lookup map of sites for enriching campaign data
  const siteMap = useMemo(() => {
    const map = new Map();
    sites.forEach(s => {
      const code = String(s.site_code || '').trim().toUpperCase();
      if (code) map.set(code, s);
    });
    return map;
  }, [sites]);

  // Extract top client suggestions for one-click filter chips
  const clientSuggestions = useMemo(() => {
    const counts = new Map();
    campaigns.forEach(c => {
      const name = String(c.client || c.client_name || '').trim();
      if (name && !isVacantClient(name)) {
        counts.set(name, (counts.get(name) || 0) + 1);
      }
    });
    return Array.from(counts.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 8)
      .map(([name]) => name);
  }, [campaigns]);

  // Enrich campaigns with site metadata & computed status
  const enrichedCampaigns = useMemo(() => {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    return campaigns.map(c => {
      const siteCodeUpper = String(c.site_code || '').trim().toUpperCase();
      const site = siteMap.get(siteCodeUpper);

      // Dates parsing
      const rawStart = c.start_date || c.booking_date;
      const rawEnd = c.end_date;
      const cStart = rawStart ? new Date(rawStart) : null;
      const cEnd = rawEnd ? new Date(rawEnd) : null;

      let computedStatus = 'live';
      if (cEnd && !isNaN(cEnd.getTime())) {
        const endDay = new Date(cEnd);
        endDay.setHours(23, 59, 59, 999);
        if (endDay < today) {
          computedStatus = 'completed';
        } else if (cStart && !isNaN(cStart.getTime())) {
          const startDay = new Date(cStart);
          startDay.setHours(0, 0, 0, 0);
          if (startDay > today) {
            computedStatus = 'upcoming';
          }
        }
      }

      // Calculate days tenure
      let tenureDays = c.days || '';
      if (!tenureDays && cStart && cEnd && !isNaN(cStart.getTime()) && !isNaN(cEnd.getTime())) {
        tenureDays = Math.max(1, Math.round((cEnd - cStart) / 86400000) + 1);
      }

      return {
        ...c,
        site_code: c.site_code || site?.site_code || '—',
        location: c.location || site?.address || site?.area || '—',
        city: site?.city || 'Ahmedabad',
        size: c.size || (c.width && c.height ? `${c.width}x${c.height} ft` : (site?.size || '—')),
        type: c.type || site?.type || 'Hoarding',
        width: c.width ?? site?.width ?? '',
        height: c.height ?? site?.height ?? '',
        computedStatus,
        tenureDays,
        cStart,
        cEnd
      };
    });
  }, [campaigns, siteMap]);

  // Filtered dataset matching user criteria
  const filteredList = useMemo(() => {
    const q = clientOrDisplay.trim().toLowerCase();
    const siteQ = siteFilter.trim().toLowerCase();
    const sFilter = startDate ? new Date(startDate + 'T00:00:00') : null;
    const eFilter = endDate ? new Date(endDate + 'T23:59:59') : null;

    return enrichedCampaigns.filter(c => {
      // 1. Client / Display search (also searches site code and location)
      if (q) {
        const clientText = String(c.client || c.client_name || '').toLowerCase();
        const displayText = String(c.display || c.campaign_name || c.brand || '').toLowerCase();
        const siteText = String(c.site_code || '').toLowerCase();
        const locText = String(c.location || '').toLowerCase();

        const matches = clientText.includes(q) || displayText.includes(q) || siteText.includes(q) || locText.includes(q);
        if (!matches) return false;
      }

      // 2. Specific Site Code filter
      if (siteQ) {
        if (!String(c.site_code || '').toLowerCase().includes(siteQ)) return false;
      }

      // 3. Date Range Filter: checks if booking period overlaps with [startDate, endDate]
      if (sFilter || eFilter) {
        const bStart = c.cStart;
        const bEnd = c.cEnd || c.cStart;

        if (!bStart && !bEnd) return false;

        const startCheck = bStart ? new Date(bStart) : new Date(bEnd);
        const endCheck = bEnd ? new Date(bEnd) : new Date(bStart);
        startCheck.setHours(0, 0, 0, 0);
        endCheck.setHours(23, 59, 59, 999);

        if (sFilter && eFilter) {
          // Booking must be active or overlap within the chosen range
          if (startCheck > eFilter || endCheck < sFilter) return false;
        } else if (sFilter) {
          if (endCheck < sFilter) return false;
        } else if (eFilter) {
          if (startCheck > eFilter) return false;
        }
      }

      // 4. Status Filter (Live, Upcoming, Completed)
      if (statusFilter !== 'ALL') {
        if (c.computedStatus !== statusFilter) return false;
      }

      return true;
    });
  }, [enrichedCampaigns, clientOrDisplay, siteFilter, startDate, endDate, statusFilter]);

  // Sorted list
  const sortedList = useMemo(() => {
    return [...filteredList].sort((a, b) => {
      let valA = a[sortState.key];
      let valB = b[sortState.key];

      if (['advt_fees', 'printing_mounting_cost', 'total_amount', 'pending', 'days', 'tenureDays'].includes(sortState.key)) {
        valA = Number(valA || 0);
        valB = Number(valB || 0);
      } else if (['start_date', 'end_date'].includes(sortState.key)) {
        valA = valA ? new Date(valA).getTime() : 0;
        valB = valB ? new Date(valB).getTime() : 0;
      } else {
        valA = String(valA || '').toLowerCase();
        valB = String(valB || '').toLowerCase();
      }

      if (valA < valB) return sortState.dir === 'asc' ? -1 : 1;
      if (valA > valB) return sortState.dir === 'asc' ? 1 : -1;
      return 0;
    });
  }, [filteredList, sortState]);

  // Summary KPIs for current filter
  const kpis = useMemo(() => {
    const totalBookings = filteredList.length;
    const uniqueSites = new Set(filteredList.map(c => c.site_code).filter(c => c && c !== '—')).size;
    let totalRevenue = 0;
    let totalAdvtFees = 0;
    let totalPending = 0;
    let totalDays = 0;

    filteredList.forEach(c => {
      totalRevenue += Number(c.total_amount || c.revenue || 0);
      totalAdvtFees += Number(c.advt_fees || 0);
      totalPending += Number(c.pending || 0);
      totalDays += Number(c.tenureDays || c.days || 0);
    });

    return { totalBookings, uniqueSites, totalRevenue, totalAdvtFees, totalPending, totalDays };
  }, [filteredList]);

  // Preset Date Range buttons
  function applyDatePreset(preset) {
    const today = new Date();
    const y = today.getFullYear();
    const m = today.getMonth();

    if (preset === 'all') {
      setStartDate('');
      setEndDate('');
    } else if (preset === 'thisMonth') {
      const firstDay = new Date(y, m, 1).toISOString().slice(0, 10);
      const lastDay = new Date(y, m + 1, 0).toISOString().slice(0, 10);
      setStartDate(firstDay);
      setEndDate(lastDay);
    } else if (preset === 'nextMonth') {
      const firstDay = new Date(y, m + 1, 1).toISOString().slice(0, 10);
      const lastDay = new Date(y, m + 2, 0).toISOString().slice(0, 10);
      setStartDate(firstDay);
      setEndDate(lastDay);
    } else if (preset === 'next30Days') {
      const start = today.toISOString().slice(0, 10);
      const end = new Date(today.getTime() + 30 * 86400000).toISOString().slice(0, 10);
      setStartDate(start);
      setEndDate(end);
    } else if (preset === 'year2026') {
      setStartDate('2026-01-01');
      setEndDate('2026-12-31');
    }
  }

  function resetAllFilters() {
    setClientOrDisplay('');
    setStartDate('');
    setEndDate('');
    setStatusFilter('ALL');
    setSiteFilter('');
  }

  function handleSort(key) {
    setSortState(prev => {
      if (prev.key === key) {
        return { key, dir: prev.dir === 'asc' ? 'desc' : 'asc' };
      }
      return { key, dir: 'asc' };
    });
  }

  // Edit / Save Campaign
  function openEdit(c) {
    setEditModal(c);
    setModalSiteCode(c.site_code || '');
    setModalClient(c.client || c.client_name || '');
    setModalDisplay(c.display || c.campaign_name || c.brand || '');
    setModalStartDate(c.start_date || c.booking_date || '');
    setModalEndDate(c.end_date || '');
    setModalAdvtFees(c.advt_fees || '');
    setModalProdCost(c.printing_mounting_cost || c.printing_cost || '');
    setModalTotalAmount(c.total_amount || c.revenue || '');
    setModalPending(c.pending || '');
    setModalNotes(c.notes || '');
  }

  async function handleSaveEdit(e) {
    e.preventDefault();
    if (!editModal) return;
    setSaving(true);
    try {
      const payload = {
        site_code: modalSiteCode,
        client: modalClient,
        client_name: modalClient,
        display: modalDisplay,
        campaign_name: modalDisplay,
        start_date: modalStartDate || null,
        end_date: modalEndDate || null,
        advt_fees: Number(modalAdvtFees) || 0,
        printing_mounting_cost: Number(modalProdCost) || 0,
        total_amount: Number(modalTotalAmount) || 0,
        pending: Number(modalPending) || 0,
        notes: modalNotes
      };

      if (editModal.id) {
        await api.put(`/campaigns/${editModal.id}`, payload);
      } else {
        await api.post('/campaigns', payload);
      }
      setEditModal(null);
      await loadData();
    } catch (err) {
      console.error('Failed to save campaign:', err);
      alert('Error saving campaign: ' + (err.response?.data?.message || err.message));
    } finally {
      setSaving(false);
    }
  }

  async function handleDeleteCampaign(id, clientName) {
    if (!canDelete) return;
    if (!confirm(`Are you sure you want to delete this booking for "${clientName || 'this client'}"?`)) {
      return;
    }
    try {
      await api.delete(`/campaigns/${id}`);
      await loadData();
    } catch (err) {
      console.error('Failed to delete campaign:', err);
      alert('Error deleting campaign.');
    }
  }

  const isFiltered = clientOrDisplay || startDate || endDate || statusFilter !== 'ALL' || siteFilter;

  return (
    <>
      <div className="scooh-pagehead">
        <div>
          <div style={{ fontSize: '11px', fontWeight: 800, color: '#a78bfa', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: '4px' }}>
            Media Buzz — OOH Workspace
          </div>
          <h1 style={{ fontSize: '26px', fontWeight: 900, margin: 0, color: '#fff', letterSpacing: '-0.02em' }}>
            Campaign Details
          </h1>
          <p style={{ margin: '6px 0 0', color: '#94a3b8', fontSize: '13px' }}>
            Search by client name, display brand, and filter by start & end date to view all campaign bookings across your sites.
          </p>
        </div>
        <div className="scooh-headactions" style={{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap' }}>
          <button type="button" className="scooh-btn ghost" onClick={loadData}>
            🔄 Refresh
          </button>
          <button
            type="button"
            className="scooh-btn ghost"
            onClick={() => exportCampaignDetailsExcel(sortedList, { clientOrDisplay, startDate, endDate })}
            title="Export filtered campaign details to Excel"
          >
            📥 Export Excel
          </button>
          {canAdd && (
            <label
              className="scooh-btn ghost"
              style={{ cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: '6px' }}
              title="Upload Excel spreadsheet (.xlsx, .xls) to import campaign bookings"
            >
              <span>📁 {importingExcel ? 'Importing…' : 'Import Excel'}</span>
              <input
                type="file"
                accept=".xlsx,.xls"
                hidden
                disabled={importingExcel}
                onChange={handleImportExcel}
              />
            </label>
          )}
          <button
            type="button"
            className="scooh-btn ghost"
            onClick={downloadSampleCampaignExcel}
            style={{ fontSize: '12px', color: '#94a3b8' }}
            title="Download sample Excel template for importing campaign bookings"
          >
            📄 Sample Template
          </button>
          {canAdd && (
            <button
              type="button"
              className="scooh-btn purple-btn"
              onClick={() => {
                setEditModal({});
                setModalSiteCode('');
                setModalClient('');
                setModalDisplay('');
                setModalStartDate(new Date().toISOString().slice(0, 10));
                setModalEndDate('');
                setModalAdvtFees('');
                setModalProdCost('');
                setModalTotalAmount('');
                setModalPending('');
                setModalNotes('');
              }}
            >
              + Add Campaign
            </button>
          )}
        </div>
      </div>

      {importBanner && (
        <div style={{ padding: '10px 18px', background: 'rgba(16, 185, 129, 0.15)', border: '1px solid rgba(16, 185, 129, 0.4)', borderRadius: '8px', color: '#34d399', fontSize: '13px', fontWeight: 600, marginBottom: '14px', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <span>{importBanner}</span>
          <button type="button" onClick={() => setImportBanner('')} style={{ background: 'none', border: 'none', color: '#34d399', cursor: 'pointer', fontSize: '16px', lineHeight: 1 }}>×</button>
        </div>
      )}

      {/* ── Filter Control Panel ────────────────────────────────────────── */}
      <section className="scooh-panel" style={{ marginBottom: '16px', padding: '18px 20px', border: '1px solid rgba(168, 85, 247, 0.25)', background: 'linear-gradient(180deg, rgba(15, 23, 42, 0.85) 0%, rgba(11, 16, 22, 0.95) 100%)' }}>
        
        {/* Main 3-Column Search & Date Pickers */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: '16px', alignItems: 'flex-end' }}>
          
          {/* Input 1: Client Name or Display / Campaign */}
          <div>
            <label style={{ display: 'block', fontSize: '12px', fontWeight: 800, color: '#f8fafc', marginBottom: '6px', letterSpacing: '0.02em' }}>
              👤 Client Name / Display Brand
            </label>
            <div style={{ position: 'relative' }}>
              <input
                type="text"
                className="scooh-search"
                style={{ width: '100%', padding: '10px 36px 10px 14px', fontSize: '13px', background: '#0b1016', border: '1px solid rgba(168, 85, 247, 0.4)', borderRadius: '8px', color: '#fff' }}
                placeholder="Type Client or Display (e.g. Tata, HDFC, EV)…"
                value={clientOrDisplay}
                onChange={e => setClientOrDisplay(e.target.value)}
              />
              {clientOrDisplay && (
                <button
                  type="button"
                  onClick={() => setClientOrDisplay('')}
                  style={{ position: 'absolute', right: '10px', top: '50%', transform: 'translateY(-50%)', background: 'transparent', border: 0, color: '#94a3b8', cursor: 'pointer', fontSize: '14px', fontWeight: 800 }}
                  title="Clear search"
                >
                  ✕
                </button>
              )}
            </div>
          </div>

          {/* Input 2: Start Date */}
          <div>
            <label style={{ display: 'block', fontSize: '12px', fontWeight: 800, color: '#f8fafc', marginBottom: '6px' }}>
              📅 Start Date (From)
            </label>
            <input
              type="date"
              className="scooh-input"
              style={{ width: '100%', padding: '9px 12px', fontSize: '13px', background: '#0b1016', border: '1px solid rgba(56, 189, 248, 0.4)', borderRadius: '8px', color: '#fff' }}
              value={startDate}
              onChange={e => setStartDate(e.target.value)}
            />
          </div>

          {/* Input 3: End Date */}
          <div>
            <label style={{ display: 'block', fontSize: '12px', fontWeight: 800, color: '#f8fafc', marginBottom: '6px' }}>
              🏁 End Date (To)
            </label>
            <input
              type="date"
              className="scooh-input"
              style={{ width: '100%', padding: '9px 12px', fontSize: '13px', background: '#0b1016', border: '1px solid rgba(56, 189, 248, 0.4)', borderRadius: '8px', color: '#fff' }}
              value={endDate}
              onChange={e => setEndDate(e.target.value)}
            />
          </div>

          {/* Input 4: Quick Status Pills & Reset */}
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center', flexWrap: 'wrap' }}>
            <div style={{ display: 'inline-flex', background: '#070b10', padding: '3px', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.08)' }}>
              {[
                { id: 'ALL', label: 'All Status' },
                { id: 'live', label: '🟢 Live' },
                { id: 'upcoming', label: '🔵 Upcoming' },
                { id: 'completed', label: '⚪ Ended' }
              ].map(s => (
                <button
                  key={s.id}
                  type="button"
                  onClick={() => setStatusFilter(s.id)}
                  style={{
                    padding: '6px 11px',
                    fontSize: '11.5px',
                    fontWeight: 700,
                    borderRadius: '6px',
                    border: 0,
                    cursor: 'pointer',
                    background: statusFilter === s.id ? '#7c3aed' : 'transparent',
                    color: statusFilter === s.id ? '#fff' : '#94a3b8',
                    transition: 'all 0.15s ease'
                  }}
                >
                  {s.label}
                </button>
              ))}
            </div>

            {isFiltered && (
              <button
                type="button"
                className="scooh-btn ghost"
                style={{ padding: '7px 12px', fontSize: '11.5px', color: '#f87171', borderColor: 'rgba(248, 113, 113, 0.4)' }}
                onClick={resetAllFilters}
                title="Reset all filters"
              >
                ✕ Clear Filters
              </button>
            )}
          </div>
        </div>

        {/* Date Presets Toolbar & Quick Client Chips */}
        <div style={{ marginTop: '14px', paddingTop: '12px', borderTop: '1px solid rgba(255,255,255,0.06)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '12px', flexWrap: 'wrap' }}>
          
          {/* Quick Date Presets */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>
            <span style={{ fontSize: '11.5px', fontWeight: 800, color: '#64748b', textTransform: 'uppercase', marginRight: '4px' }}>
              Date Presets:
            </span>
            <button type="button" className="scooh-pill" onClick={() => applyDatePreset('all')} style={{ cursor: 'pointer', padding: '3px 9px', fontSize: '11px', background: !startDate && !endDate ? 'rgba(56, 189, 248, 0.2)' : 'rgba(255,255,255,0.04)', color: !startDate && !endDate ? '#38bdf8' : '#94a3b8' }}>
              All Time
            </button>
            <button type="button" className="scooh-pill" onClick={() => applyDatePreset('thisMonth')} style={{ cursor: 'pointer', padding: '3px 9px', fontSize: '11px', background: 'rgba(255,255,255,0.04)', color: '#94a3b8' }}>
              This Month
            </button>
            <button type="button" className="scooh-pill" onClick={() => applyDatePreset('nextMonth')} style={{ cursor: 'pointer', padding: '3px 9px', fontSize: '11px', background: 'rgba(255,255,255,0.04)', color: '#94a3b8' }}>
              Next Month
            </button>
            <button type="button" className="scooh-pill" onClick={() => applyDatePreset('next30Days')} style={{ cursor: 'pointer', padding: '3px 9px', fontSize: '11px', background: 'rgba(255,255,255,0.04)', color: '#94a3b8' }}>
              Next 30 Days
            </button>
            <button type="button" className="scooh-pill" onClick={() => applyDatePreset('year2026')} style={{ cursor: 'pointer', padding: '3px 9px', fontSize: '11px', background: 'rgba(255,255,255,0.04)', color: '#94a3b8' }}>
              Year 2026
            </button>
          </div>

          {/* Quick Filter by Client Chips */}
          {clientSuggestions.length > 0 && (
            <div style={{ display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>
              <span style={{ fontSize: '11.5px', fontWeight: 800, color: '#64748b', textTransform: 'uppercase', marginRight: '4px' }}>
                Top Clients:
              </span>
              {clientSuggestions.slice(0, 5).map(cName => (
                <button
                  key={cName}
                  type="button"
                  className="scooh-pill"
                  onClick={() => setClientOrDisplay(cName)}
                  style={{
                    cursor: 'pointer',
                    padding: '2px 8px',
                    fontSize: '11px',
                    background: clientOrDisplay === cName ? 'rgba(168, 85, 247, 0.25)' : 'rgba(255,255,255,0.04)',
                    color: clientOrDisplay === cName ? '#c084fc' : '#cbd5e1',
                    border: clientOrDisplay === cName ? '1px solid rgba(168, 85, 247, 0.5)' : '1px solid transparent'
                  }}
                  title={`Filter bookings for ${cName}`}
                >
                  {cName}
                </button>
              ))}
            </div>
          )}
        </div>
      </section>

      {/* ── Matching Bookings Table ──────────────────────────────────────── */}
      <div className="scooh-panel" style={{ overflow: 'hidden' }}>
        <div style={{ padding: '14px 20px', borderBottom: '1px solid rgba(255,255,255,0.06)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px' }}>
          <div>
            <span style={{ fontSize: '14px', fontWeight: 800, color: '#f8fafc' }}>
              Bookings Ledger ({sortedList.length})
            </span>
            {isFiltered && (
              <span style={{ fontSize: '12px', color: '#a78bfa', marginLeft: '10px' }}>
                Filtered by: {clientOrDisplay ? `"${clientOrDisplay}" ` : ''}{startDate ? `from ${startDate} ` : ''}{endDate ? `to ${endDate} ` : ''}
              </span>
            )}
          </div>
          <div style={{ fontSize: '12px', color: '#64748b' }}>
            Click column headers to sort • Click on any booking to inspect full details
          </div>
        </div>

        <div className="scooh-tablewrap" style={{ overflowX: 'auto' }}>
          <table className="scooh-table" style={{ minWidth: '1050px' }}>
            <thead>
              <tr>
                <th style={{ width: '45px', textAlign: 'center' }}>#</th>
                <th style={{ minWidth: '105px', cursor: 'pointer' }} onClick={() => handleSort('site_code')}>
                  Site Code {sortState.key === 'site_code' && (sortState.dir === 'asc' ? '↑' : '↓')}
                </th>
                <th style={{ minWidth: '220px', cursor: 'pointer' }} onClick={() => handleSort('client')}>
                  Client / Display {sortState.key === 'client' && (sortState.dir === 'asc' ? '↑' : '↓')}
                </th>
                <th style={{ minWidth: '220px' }}>Location</th>
                <th style={{ minWidth: '120px', textAlign: 'center' }}>Size & Type</th>
                <th style={{ minWidth: '110px', textAlign: 'center', cursor: 'pointer' }} onClick={() => handleSort('start_date')}>
                  Start Date {sortState.key === 'start_date' && (sortState.dir === 'asc' ? '↑' : '↓')}
                </th>
                <th style={{ minWidth: '110px', textAlign: 'center', cursor: 'pointer' }} onClick={() => handleSort('end_date')}>
                  End Date {sortState.key === 'end_date' && (sortState.dir === 'asc' ? '↑' : '↓')}
                </th>
                <th style={{ minWidth: '75px', textAlign: 'center', cursor: 'pointer' }} onClick={() => handleSort('days')}>
                  Days {sortState.key === 'days' && (sortState.dir === 'asc' ? '↑' : '↓')}
                </th>
                
                <th style={{ minWidth: '110px', textAlign: 'center' }}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {loading ? (
                <tr>
                  <td colSpan={9} style={{ textAlign: 'center', padding: '40px 20px', color: '#94a3b8' }}>
                    <div style={{ fontSize: '18px', marginBottom: '8px' }}>⏳ Loading campaign details…</div>
                  </td>
                </tr>
              ) : sortedList.length === 0 ? (
                <tr>
                  <td colSpan={9} style={{ textAlign: 'center', padding: '50px 20px', color: '#64748b' }}>
                    <div style={{ fontSize: '40px', marginBottom: '10px' }}>🔍</div>
                    <div style={{ fontSize: '16px', fontWeight: 700, color: '#cbd5e1' }}>
                      No campaigns found for the selected filter
                    </div>
                    <div style={{ fontSize: '13px', color: '#94a3b8', marginTop: '6px', maxWidth: '420px', margin: '6px auto 16px' }}>
                      {clientOrDisplay || startDate || endDate
                        ? `No bookings match "${clientOrDisplay || 'any client'}" within ${startDate || 'any'} to ${endDate || 'any'}.`
                        : 'No campaign records exist yet.'}
                    </div>
                    {isFiltered && (
                      <button type="button" className="scooh-btn purple-btn" onClick={resetAllFilters}>
                        Clear All Filters
                      </button>
                    )}
                  </td>
                </tr>
              ) : (
                sortedList.map((c, idx) => {
                  const isPending = Number(c.pending || 0) > 0;
                  const isLive = c.computedStatus === 'live';
                  const isUpcoming = c.computedStatus === 'upcoming';
                  const statusColor = isLive ? '#4ade80' : (isUpcoming ? '#38bdf8' : '#94a3b8');
                  const statusBg = isLive ? 'rgba(74, 222, 128, 0.12)' : (isUpcoming ? 'rgba(56, 189, 248, 0.12)' : 'rgba(148, 163, 184, 0.1)');
                  const statusBorder = isLive ? 'rgba(74, 222, 128, 0.3)' : (isUpcoming ? 'rgba(56, 189, 248, 0.3)' : 'rgba(148, 163, 184, 0.2)');

                  return (
                    <tr
                      key={c.id || `${c.site_code}-${idx}`}
                      style={{ cursor: 'pointer', transition: 'background 0.15s ease' }}
                      onClick={() => setSelectedCampaign(c)}
                    >
                      <td style={{ textAlign: 'center', color: '#64748b', fontSize: '11.5px' }}>
                        {idx + 1}
                      </td>
                      <td>
                        <button
                          type="button"
                          className="scooh-plate"
                          style={{ cursor: 'pointer', background: 'rgba(56, 189, 248, 0.12)', border: '1px solid rgba(56, 189, 248, 0.45)', color: '#38bdf8', padding: '3px 8px', borderRadius: '5px', fontWeight: 800 }}
                          onClick={e => {
                            e.stopPropagation();
                            navigate(`/sites?search=${encodeURIComponent(c.site_code)}`);
                          }}
                          title={`View ${c.site_code} in Sites Directory`}
                        >
                          {c.site_code}
                        </button>
                      </td>
                      <td>
                        <div style={{ fontWeight: 800, color: '#f8fafc', fontSize: '13.5px' }}>
                          {c.client || c.client_name || '—'}
                        </div>
                        {Boolean((c.display || c.campaign_name || c.brand) &&
                          String(c.display || c.campaign_name || c.brand).trim().toLowerCase() !== String(c.client || c.client_name || '').trim().toLowerCase()) && (
                          <div style={{ marginTop: '3px' }}>
                            <span style={{ display: 'inline-block', padding: '2px 7px', background: 'rgba(168, 85, 247, 0.15)', border: '1px solid rgba(168, 85, 247, 0.35)', color: '#c084fc', borderRadius: '4px', fontSize: '11.5px', fontWeight: 700 }}>
                              📢 {c.display || c.campaign_name || c.brand}
                            </span>
                          </div>
                        )}
                        {c.vendor_name && (
                          <div style={{ fontSize: '10.5px', color: '#64748b', marginTop: '2px' }}>
                            Vendor: {c.vendor_name}
                          </div>
                        )}
                      </td>
                      <td style={{ fontSize: '12px', color: '#cbd5e1', maxWidth: '240px' }}>
                        <div style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }} title={c.location}>
                          {c.location}
                        </div>
                        <div style={{ fontSize: '11px', color: '#64748b' }}>{c.city}</div>
                      </td>
                      <td style={{ textAlign: 'center', fontSize: '12px', color: '#94a3b8' }}>
                        <div style={{ fontWeight: 700, color: '#f1f5f9' }}>{c.size}</div>
                        <div style={{ fontSize: '11px', color: '#64748b' }}>{c.type}</div>
                      </td>
                      <td style={{ textAlign: 'center', fontSize: '12px', color: '#f1f5f9', fontWeight: 600 }}>
                        {formatDate(c.start_date || c.booking_date) || '—'}
                      </td>
                      <td style={{ textAlign: 'center', fontSize: '12px', color: '#f1f5f9', fontWeight: 600 }}>
                        {formatDate(c.end_date) || '—'}
                      </td>
                      <td style={{ textAlign: 'center' }}>
                        <span style={{ fontSize: '11.5px', fontWeight: 800, color: '#cbd5e1', background: '#1e293b', padding: '2px 7px', borderRadius: '4px' }}>
                          {c.tenureDays ? `${c.tenureDays}d` : '—'}
                        </span>
                      </td>
                      
                      <td style={{ textAlign: 'center' }} onClick={e => e.stopPropagation()}>
                        <div style={{ display: 'inline-flex', gap: '5px', alignItems: 'center' }}>
                          <button
                            type="button"
                            className="scooh-btn ghost"
                            style={{ padding: '3px 7px', fontSize: '11px' }}
                            onClick={() => setSelectedCampaign(c)}
                            title="Inspect complete campaign record"
                          >
                            👁️
                          </button>
                          {canAdd && (
                            <button
                              type="button"
                              className="scooh-btn ghost"
                              style={{ padding: '3px 7px', fontSize: '11px' }}
                              onClick={() => openEdit(c)}
                              title="Edit this booking"
                            >
                              ✏️
                            </button>
                          )}
                          {canDelete && (
                            <button
                              type="button"
                              className="scooh-btn ghost"
                              style={{ padding: '3px 7px', fontSize: '11px', color: '#f87171' }}
                              onClick={() => handleDeleteCampaign(c.id, c.client || c.client_name)}
                              title="Delete booking"
                            >
                              🗑️
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

      {/* ── Modal: Full Campaign Details ─────────────────────────────────── */}
      {selectedCampaign && (
        <div className="scooh-modalwrap" onClick={() => setSelectedCampaign(null)}>
          <div className="scooh-modal" style={{ maxWidth: '640px', width: '90%' }} onClick={e => e.stopPropagation()}>
            <div className="scooh-modalhead" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
              <div>
                <div style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', marginBottom: '6px' }}>
                  <span className="scooh-plate" style={{ fontSize: '14px', fontWeight: 900, background: 'rgba(56, 189, 248, 0.15)', color: '#38bdf8' }}>
                    {selectedCampaign.site_code}
                  </span>
                  <span style={{ fontSize: '11px', fontWeight: 800, padding: '2px 8px', borderRadius: '4px', background: selectedCampaign.computedStatus === 'live' ? 'rgba(74, 222, 128, 0.15)' : 'rgba(148, 163, 184, 0.15)', color: selectedCampaign.computedStatus === 'live' ? '#4ade80' : '#94a3b8' }}>
                    {selectedCampaign.computedStatus ? selectedCampaign.computedStatus.toUpperCase() : 'ACTIVE'}
                  </span>
                </div>
                <h3 style={{ margin: 0, fontSize: '20px', color: '#fff', fontWeight: 800 }}>
                  {selectedCampaign.client || selectedCampaign.client_name || 'Client Booking'}
                </h3>
                <div style={{ color: '#c084fc', fontSize: '13px', fontWeight: 700, marginTop: '2px' }}>
                  📢 {selectedCampaign.display || selectedCampaign.campaign_name || selectedCampaign.brand || 'Campaign'}
                </div>
              </div>
              <button type="button" className="scooh-btn ghost" onClick={() => setSelectedCampaign(null)} style={{ padding: '4px 10px' }}>
                ✕
              </button>
            </div>

            <div className="scooh-modalbody" style={{ display: 'flex', flexDirection: 'column', gap: '16px', padding: '20px' }}>
              
              {/* Financial Breakdown Card */}
              <div style={{ background: '#0b1016', padding: '14px 16px', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.08)' }}>
                <div style={{ fontSize: '11px', fontWeight: 800, color: '#94a3b8', textTransform: 'uppercase', marginBottom: '8px' }}>
                  Financial Summary
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(130px, 1fr))', gap: '12px' }}>
                  <div>
                    <span style={{ fontSize: '11px', color: '#64748b' }}>Advt. Fees / Mo</span>
                    <div style={{ fontSize: '16px', fontWeight: 800, color: '#f1f5f9' }}>
                      {Number(selectedCampaign.advt_fees || 0) > 0 ? money(Number(selectedCampaign.advt_fees)) : '—'}
                    </div>
                  </div>
                  <div>
                    <span style={{ fontSize: '11px', color: '#64748b' }}>Printing & Mounting</span>
                    <div style={{ fontSize: '16px', fontWeight: 800, color: '#f1f5f9' }}>
                      {Number(selectedCampaign.printing_mounting_cost || selectedCampaign.printing_cost || 0) > 0 ? money(Number(selectedCampaign.printing_mounting_cost || selectedCampaign.printing_cost)) : '—'}
                    </div>
                  </div>
                  <div>
                    <span style={{ fontSize: '11px', color: '#64748b' }}>Total Amount</span>
                    <div style={{ fontSize: '18px', fontWeight: 900, color: '#34d399' }}>
                      {Number(selectedCampaign.total_amount || selectedCampaign.revenue || 0) > 0 ? money(Number(selectedCampaign.total_amount || selectedCampaign.revenue)) : '—'}
                    </div>
                  </div>

                </div>
              </div>

              {/* Schedule & Dates */}
              <div style={{ background: '#0b1016', padding: '14px 16px', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.08)' }}>
                <div style={{ fontSize: '11px', fontWeight: 800, color: '#94a3b8', textTransform: 'uppercase', marginBottom: '8px' }}>
                  Campaign Schedule
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(130px, 1fr))', gap: '12px' }}>
                  <div>
                    <span style={{ fontSize: '11px', color: '#64748b' }}>Start Date</span>
                    <div style={{ fontSize: '14px', fontWeight: 800, color: '#f1f5f9' }}>
                      {formatDate(selectedCampaign.start_date || selectedCampaign.booking_date) || '—'}
                    </div>
                  </div>
                  <div>
                    <span style={{ fontSize: '11px', color: '#64748b' }}>End Date</span>
                    <div style={{ fontSize: '14px', fontWeight: 800, color: '#f1f5f9' }}>
                      {formatDate(selectedCampaign.end_date) || '—'}
                    </div>
                  </div>
                  <div>
                    <span style={{ fontSize: '11px', color: '#64748b' }}>Duration</span>
                    <div style={{ fontSize: '14px', fontWeight: 800, color: '#c084fc' }}>
                      {selectedCampaign.tenureDays ? `${selectedCampaign.tenureDays} Days` : '—'}
                    </div>
                  </div>
                  {selectedCampaign.booking_date && (
                    <div>
                      <span style={{ fontSize: '11px', color: '#64748b' }}>Booked On</span>
                      <div style={{ fontSize: '14px', fontWeight: 800, color: '#f1f5f9' }}>
                        {formatDate(selectedCampaign.booking_date)}
                      </div>
                    </div>
                  )}
                </div>
              </div>

              {/* Site Specifications */}
              <div style={{ background: '#0b1016', padding: '14px 16px', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.08)' }}>
                <div style={{ fontSize: '11px', fontWeight: 800, color: '#94a3b8', textTransform: 'uppercase', marginBottom: '8px' }}>
                  Site & Hardware Details
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', fontSize: '13px' }}>
                  <div>
                    <span style={{ color: '#64748b' }}>Location: </span>
                    <span style={{ color: '#f1f5f9', fontWeight: 700 }}>{selectedCampaign.location}</span>
                  </div>
                  <div>
                    <span style={{ color: '#64748b' }}>Dimensions: </span>
                    <span style={{ color: '#f1f5f9', fontWeight: 700 }}>{selectedCampaign.size} • {selectedCampaign.type}</span>
                  </div>
                  {selectedCampaign.vendor_name && (
                    <div>
                      <span style={{ color: '#64748b' }}>Vendor: </span>
                      <span style={{ color: '#38bdf8', fontWeight: 700 }}>{selectedCampaign.vendor_name}</span>
                    </div>
                  )}
                  {selectedCampaign.notes && (
                    <div style={{ marginTop: '6px', padding: '8px', background: '#1e293b', borderRadius: '6px', color: '#cbd5e1', fontSize: '12px' }}>
                      📝 <b>Notes:</b> {selectedCampaign.notes}
                    </div>
                  )}
                </div>
              </div>
            </div>

            <div className="scooh-modalfoot" style={{ display: 'flex', justifyContent: 'space-between', padding: '14px 20px' }}>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={() => {
                  const site = selectedCampaign.site_code;
                  setSelectedCampaign(null);
                  navigate(`/sites?search=${encodeURIComponent(site)}`);
                }}
              >
                📍 View Site Details
              </button>
              <div style={{ display: 'flex', gap: '10px' }}>
                {canAdd && (
                  <button
                    type="button"
                    className="scooh-btn ghost"
                    onClick={() => {
                      const c = selectedCampaign;
                      setSelectedCampaign(null);
                      openEdit(c);
                    }}
                  >
                    ✏️ Edit Booking
                  </button>
                )}
                <button type="button" className="scooh-btn purple-btn" onClick={() => setSelectedCampaign(null)}>
                  Close
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ── Modal: Add / Edit Campaign ───────────────────────────────────── */}
      {editModal && (
        <div className="scooh-modalwrap" onClick={() => setEditModal(null)}>
          <div className="scooh-modal" style={{ maxWidth: '580px', width: '90%' }} onClick={e => e.stopPropagation()}>
            <form onSubmit={handleSaveEdit}>
              <div className="scooh-modalhead">
                <h3 style={{ margin: 0, fontSize: '18px', color: '#fff' }}>
                  {editModal.id ? 'Edit Campaign Booking' : 'Add New Campaign Booking'}
                </h3>
              </div>

              <div className="scooh-modalbody" style={{ display: 'flex', flexDirection: 'column', gap: '14px', padding: '18px' }}>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                  <div>
                    <label style={{ fontSize: '12px', fontWeight: 800, color: '#cbd5e1' }}>Site Code *</label>
                    <input
                      type="text"
                      required
                      className="scooh-input"
                      style={{ width: '100%', marginTop: '4px' }}
                      value={modalSiteCode}
                      onChange={e => setModalSiteCode(e.target.value.toUpperCase())}
                      placeholder="e.g. MB-01"
                    />
                  </div>
                  <div>
                    <label style={{ fontSize: '12px', fontWeight: 800, color: '#cbd5e1' }}>Client / Agency Name *</label>
                    <input
                      type="text"
                      required
                      className="scooh-input"
                      style={{ width: '100%', marginTop: '4px' }}
                      value={modalClient}
                      onChange={e => setModalClient(e.target.value)}
                      placeholder="e.g. Tata Motors"
                    />
                  </div>
                </div>

                <div>
                  <label style={{ fontSize: '12px', fontWeight: 800, color: '#cbd5e1' }}>Display / Campaign Brand *</label>
                  <input
                    type="text"
                    required
                    className="scooh-input"
                    style={{ width: '100%', marginTop: '4px' }}
                    value={modalDisplay}
                    onChange={e => setModalDisplay(e.target.value)}
                    placeholder="e.g. Tata Curvv EV Launch"
                  />
                </div>

                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                  <div>
                    <label style={{ fontSize: '12px', fontWeight: 800, color: '#cbd5e1' }}>Start Date *</label>
                    <input
                      type="date"
                      required
                      className="scooh-input"
                      style={{ width: '100%', marginTop: '4px' }}
                      value={modalStartDate}
                      onChange={e => setModalStartDate(e.target.value)}
                    />
                  </div>
                  <div>
                    <label style={{ fontSize: '12px', fontWeight: 800, color: '#cbd5e1' }}>End Date *</label>
                    <input
                      type="date"
                      required
                      className="scooh-input"
                      style={{ width: '100%', marginTop: '4px' }}
                      value={modalEndDate}
                      onChange={e => setModalEndDate(e.target.value)}
                    />
                  </div>
                </div>

                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                  <div>
                    <label style={{ fontSize: '12px', fontWeight: 800, color: '#cbd5e1' }}>Advt. Fees / Month (₹)</label>
                    <input
                      type="number"
                      className="scooh-input"
                      style={{ width: '100%', marginTop: '4px' }}
                      value={modalAdvtFees}
                      onChange={e => setModalAdvtFees(e.target.value)}
                      placeholder="e.g. 250000"
                    />
                  </div>
                  <div>
                    <label style={{ fontSize: '12px', fontWeight: 800, color: '#cbd5e1' }}>Printing & Mounting (₹)</label>
                    <input
                      type="number"
                      className="scooh-input"
                      style={{ width: '100%', marginTop: '4px' }}
                      value={modalProdCost}
                      onChange={e => setModalProdCost(e.target.value)}
                      placeholder="e.g. 20000"
                    />
                  </div>
                </div>

                <div>
                  <label style={{ fontSize: '12px', fontWeight: 800, color: '#cbd5e1' }}>Total Gross Amount (₹)</label>
                  <input
                    type="number"
                    className="scooh-input"
                    style={{ width: '100%', marginTop: '4px' }}
                    value={modalTotalAmount}
                    onChange={e => setModalTotalAmount(e.target.value)}
                    placeholder="Auto sum or enter"
                  />
                </div>

                <div>
                  <label style={{ fontSize: '12px', fontWeight: 800, color: '#cbd5e1' }}>Notes</label>
                  <textarea
                    className="scooh-input"
                    rows={2}
                    style={{ width: '100%', marginTop: '4px' }}
                    value={modalNotes}
                    onChange={e => setModalNotes(e.target.value)}
                    placeholder="Additional notes or booking remarks…"
                  />
                </div>
              </div>

              <div className="scooh-modalfoot" style={{ display: 'flex', justifyContent: 'flex-end', gap: '10px' }}>
                <button type="button" className="scooh-btn ghost" onClick={() => setEditModal(null)}>
                  Cancel
                </button>
                <button type="submit" className="scooh-btn purple-btn" disabled={saving}>
                  {saving ? 'Saving…' : 'Save Campaign'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </>
  );
}

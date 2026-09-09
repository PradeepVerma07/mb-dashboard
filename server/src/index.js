import 'dotenv/config';
import express from 'express';import cors from 'cors';import helmet from 'helmet';import path from 'path';import fs from 'fs';import multer from 'multer';import bcrypt from 'bcryptjs';import {fileURLToPath} from 'url';import * as xlsxModule from 'xlsx';
const XLSX = xlsxModule.default || xlsxModule;
import {pool,q} from './db.js';import {auth,admin,sign,managerOrAdmin,notViewer,requireRole} from './auth.js';
import {computeAllOccupied, canonicalSiteCode, getConflictSummary, syncLinkedCampaigns, deleteLinkedCampaigns, batchDeleteLinkedCampaigns, syncAllLinkedCampaigns} from './siteHierarchy.js';
const __dirname=path.dirname(fileURLToPath(import.meta.url));const app=express();const PORT=Number(process.env.PORT||3000);const uploadDir=path.resolve(__dirname,'..',process.env.UPLOAD_DIR||'uploads');fs.mkdirSync(uploadDir,{recursive:true});
app.use(helmet({crossOriginResourcePolicy:false}));app.use(cors({origin:true,credentials:true}));app.use(express.json({limit:'10mb'}));app.use('/uploads',express.static(uploadDir));
const upload=multer({dest:uploadDir,limits:{fileSize:Number(process.env.MAX_UPLOAD_MB||20)*1024*1024}});
const entities={
 sites:{table:'sites',order:'id DESC'},clients:{table:'clients',order:'id DESC'},campaigns:{table:'campaigns',order:'id DESC'},validations:{table:'validations',order:'id DESC'},electricity:{table:'electricity',order:'id DESC'},vendors:{table:'vendors',order:'id DESC'},'vendor-jobs':{table:'vendor_jobs',order:'id DESC'},proposals:{table:'proposals',order:'id DESC'},invoices:{table:'invoices',order:'id DESC'}
};
const blocked=new Set(['id','created_at','updated_at']);
function safeEntity(req,res){const e=entities[req.params.entity];if(!e){res.status(404).json({message:'Unknown module'});return null}return e}
async function tableColumns(table){const rows=await q(`SHOW COLUMNS FROM \`${table}\``);return new Set(rows.map(r=>r.Field))}
async function cleanData(table,body){const cols=await tableColumns(table),out={};for(const [k,v] of Object.entries(body||{})){if(cols.has(k)&&!blocked.has(k)){if(typeof v==='object'&&v!==null)out[k]=JSON.stringify(v);else out[k]=v===''?null:v}}return out}
async function audit(user,action,type,id,before=null,after=null){try{await q('INSERT INTO activity_log (user_id,user_name,action,object_type,object_id,old_value,new_value,created_at) VALUES (?,?,?,?,?,?,?,NOW())',[user?.id||null,user?.name||'',action,type,id||null,before?JSON.stringify(before):null,after?JSON.stringify(after):null])}catch{}}

export async function ensureCampaignColumns() {
  try {
    const existingCols = await tableColumns('campaigns');
    const colsToAdd = [
      { name: 'month', type: 'VARCHAR(60) NOT NULL DEFAULT ""' },
      { name: 'display', type: 'VARCHAR(190) NOT NULL DEFAULT ""' },
      { name: 'vendor_name', type: 'VARCHAR(190) NOT NULL DEFAULT ""' },
      { name: 'location', type: 'VARCHAR(255) NOT NULL DEFAULT ""' },
      { name: 'width', type: 'DECIMAL(10,2) NULL' },
      { name: 'height', type: 'DECIMAL(10,2) NULL' },
      { name: 'size', type: 'VARCHAR(60) NOT NULL DEFAULT ""' },
      { name: 'type', type: 'VARCHAR(80) NOT NULL DEFAULT "Hoarding"' },
      { name: 'days', type: 'INT NOT NULL DEFAULT 30' },
      { name: 'advt_fees', type: 'DECIMAL(15,2) NOT NULL DEFAULT 0' },
      { name: 'printing_mounting_cost', type: 'DECIMAL(15,2) NOT NULL DEFAULT 0' },
      { name: 'total_amount', type: 'DECIMAL(15,2) NOT NULL DEFAULT 0' },
      { name: 'po', type: 'VARCHAR(100) NOT NULL DEFAULT ""' },
      { name: 'bill', type: 'VARCHAR(100) NOT NULL DEFAULT ""' },
      { name: 'pending', type: 'DECIMAL(15,2) NOT NULL DEFAULT 0' },
      { name: 'occupancy', type: 'VARCHAR(100) NOT NULL DEFAULT ""' }
    ];
    for (const c of colsToAdd) {
      if (!existingCols.has(c.name)) {
        try {
          await q(`ALTER TABLE campaigns ADD COLUMN \`${c.name}\` ${c.type}`);
          console.log(`[Schema Migration] Added column \`${c.name}\` to campaigns table.`);
        } catch (err) {
          console.warn(`[Schema Migration] Notice adding \`${c.name}\`:`, err.message);
        }
      }
    }
  } catch (err) {
    console.warn('[Schema Migration] ensureCampaignColumns notice:', err.message);
  }
}

/**
 * syncSiteAvailability() – Recalculates and persists availability for ALL sites
 * based on currently active campaigns. Uses the physical panel conflict map
 * so booking a combined site automatically marks all overlapping split sites
 * as Occupied, and vice-versa.
 */
async function syncSiteAvailability() {
  try {
    // 1. Get all active campaign site codes
    const activeCampaigns = await q(
      `SELECT site_code FROM campaigns WHERE record_status='active' AND site_code != '' AND (end_date IS NULL OR end_date >= CURDATE())`
    );
    const directBooked = [];
    for (const c of activeCampaigns) {
      const raw = String(c.site_code || '').trim();
      const matches = raw.match(/MB[-_ ]?\d+/gi);
      if (matches) {
        for (const m of matches) directBooked.push(canonicalSiteCode(m));
      } else if (raw) {
        directBooked.push(canonicalSiteCode(raw));
      }
    }

    // 2. Compute all occupied site codes (direct + linked via panel conflicts)
    const occupiedMap = computeAllOccupied(directBooked); // Map<siteCode, reason>

    // 3. Get all active sites
    const allSites = await q(`SELECT id, site_code, availability FROM sites WHERE record_status='active'`);

    // 4. Update each site's availability based on computed map
    for (const site of allSites) {
      const code = canonicalSiteCode(site.site_code);
      const occupiedEntry = occupiedMap.get(code);
      const currentAvail = String(site.availability || '').toLowerCase().trim();

      if (occupiedEntry) {
        // Site should be Occupied
        const reason = occupiedEntry === 'Direct' ? 'Occupied' : `Occupied (${occupiedEntry})`;
        if (currentAvail !== reason.toLowerCase() && currentAvail !== 'maintenance') {
          await q(`UPDATE sites SET availability=?, updated_at=NOW() WHERE id=?`, [reason, site.id]);
        }
      } else {
        // Site should be Available (revert only if previously marked occupied)
        if (currentAvail.startsWith('occupied')) {
          await q(`UPDATE sites SET availability='Available', updated_at=NOW() WHERE id=?`, [site.id]);
        }
      }
    }
  } catch (err) {
    console.error('[syncSiteAvailability] Error:', err.message);
  }
}
app.get('/api/health',async(req,res)=>{try{await q('SELECT 1');res.json({ok:true})}catch(e){res.status(500).json({ok:false,message:e.message})}});
app.post('/api/auth/login', async (req, res) => {
  try {
    const {email, password} = req.body || {};
    const rows = await q('SELECT * FROM users WHERE email=? AND status="active" LIMIT 1', [email]);
    const u = rows[0];
    if (!u || !await bcrypt.compare(password || '', u.password_hash)) {
      return res.status(401).json({message: 'Invalid email or password'});
    }
    res.json({token: sign(u), user: {id: u.id, email: u.email, name: u.name, role: u.role}});
  } catch (err) {
    console.error('Login error:', err.message);
    res.status(500).json({message: 'Database error: ' + err.message});
  }
});
app.get('/api/auth/me',auth,(req,res)=>res.json(req.user));
app.get('/api/dashboard', auth, async (req, res) => {
  try {
    const [
      [s],
      [cLive],
      [cAll],
      [e],
      [i],
      [p],
      campaignsRows,
      sitesRows,
      electricityRows
    ] = await Promise.all([
      // Sites count and status
      q(`SELECT 
          COUNT(*) as total, 
          COALESCE(SUM(CASE WHEN LOWER(availability) = 'available' THEN 1 ELSE 0 END), 0) as available,
          COALESCE(SUM(CASE WHEN LOWER(availability) IN ('occupied', 'booked', 'active') THEN 1 ELSE 0 END), 0) as occupied
         FROM sites WHERE record_status = 'active'`),
      // Live campaigns running right now
      q(`SELECT 
          COUNT(*) as live, 
          COALESCE(SUM(revenue), 0) as revenue, 
          COALESCE(SUM(revenue - COALESCE(vendor_cost,0) - COALESCE(printing_cost,0) - COALESCE(mounting_cost,0) - COALESCE(electricity_cost,0) - COALESCE(other_cost,0)), 0) as margin 
         FROM campaigns 
         WHERE record_status = 'active' 
           AND (
             (start_date IS NOT NULL AND end_date IS NOT NULL AND CURDATE() BETWEEN start_date AND end_date)
             OR (mounting_status = 'Mounted' AND (end_date IS NULL OR end_date >= CURDATE()))
           )`),
      // Action triggers from campaigns
      q(`SELECT 
          COUNT(*) as total, 
          COALESCE(SUM(CASE WHEN mounting_status NOT IN ('Mounted', 'Completed') AND start_date <= CURDATE() THEN 1 ELSE 0 END), 0) as mounting_overdue, 
          COALESCE(SUM(CASE WHEN validation_15_date IS NOT NULL AND validation_15_date <= CURDATE() THEN 1 ELSE 0 END), 0) as val15, 
          COALESCE(SUM(CASE WHEN final_validation_date IS NOT NULL AND final_validation_date <= CURDATE() THEN 1 ELSE 0 END), 0) as valFinal 
         FROM campaigns WHERE record_status = 'active'`),
      // Electricity bills
      q(`SELECT 
          COUNT(*) as total, 
          COALESCE(SUM(CASE WHEN payment_status != 'Paid' AND due_date < CURDATE() THEN 1 ELSE 0 END), 0) as overdue, 
          COALESCE(SUM(CASE WHEN payment_status != 'Paid' THEN 1 ELSE 0 END), 0) as pending_bills,
          COALESCE(SUM(CASE WHEN payment_status != 'Paid' THEN amount ELSE 0 END), 0) as unpaid 
         FROM electricity WHERE record_status = 'active'`),
      // Invoices
      q(`SELECT 
          COUNT(*) as total, 
          COALESCE(SUM(CASE WHEN invoice_status IN ('Pending', 'Draft') OR hard_copy_status = 'Pending' OR payment_status = 'Pending' THEN 1 ELSE 0 END), 0) as pending 
         FROM invoices WHERE record_status = 'active'`),
      // Proposals
      q(`SELECT 
          COUNT(*) as total, 
          COALESCE(SUM(CASE WHEN status IN ('Draft', 'Sent', 'Open', 'Pending') THEN 1 ELSE 0 END), 0) as open_proposals 
         FROM proposals`),
      // Active campaigns rows for live alerts
      q(`SELECT id, booking_code, site_code, client, campaign_name, start_date, end_date, mounting_status, validation_15_date, final_validation_date, invoice_status, hard_copy_status 
         FROM campaigns WHERE record_status = 'active' ORDER BY id DESC`),
      // Sites rows for classification and flags
      q(`SELECT id, site_code, area, city, availability, monthly_rate, flags FROM sites WHERE record_status = 'active'`),
      // Electricity rows for live overdue alerts
      q(`SELECT id, site_code, bill_type, amount, due_date, payment_status FROM electricity WHERE record_status = 'active' AND payment_status != 'Paid'`)
    ]);

    const totalSites = Number(s?.total || 0);

    // Active sites: sites that either have availability='occupied' OR are currently linked to a live active campaign
    const liveCampaignSiteCodes = new Set(
      (campaignsRows || [])
        .filter(c => {
          const now = new Date();
          const start = c.start_date ? new Date(c.start_date) : null;
          const end = c.end_date ? new Date(c.end_date) : null;
          return (!start || start <= now) && (!end || end >= now);
        })
        .map(c => c.site_code)
        .filter(Boolean)
    );

    let activeSites = Number(s?.occupied || 0);
    if (liveCampaignSiteCodes.size > activeSites) {
      activeSites = Math.min(totalSites, liveCampaignSiteCodes.size);
    }
    const nonActiveSites = Math.max(0, totalSites - activeSites);
    const activeCampaigns = Number(cLive?.live || 0);

    // Count flagged and premium sites
    let flaggedCount = 0;
    let primeCount = 0;
    for (const st of (sitesRows || [])) {
      try {
        const fl = typeof st.flags === 'string' ? JSON.parse(st.flags) : (st.flags || {});
        if (Array.isArray(fl) ? fl.length > 0 : (Array.isArray(fl.tags) && fl.tags.length > 0)) {
          flaggedCount++;
        }
      } catch {}
      if (Number(st.monthly_rate || 0) >= 100000) {
        primeCount++;
      }
    }

    // Real portfolio occupancy breakdown
    const occupancyBuckets = {
      'Occupied': activeSites,
      'Available': nonActiveSites,
      'Prime Sites': primeCount,
      'Needs review': flaggedCount
    };

    const avgOccupancy = totalSites > 0 ? Math.round((activeSites / totalSites) * 100) : 0;

    // Real alerts from actual active database records only
    const alerts = [];
    const now = new Date();
    for (const cmp of (campaignsRows || [])) {
      const start = cmp.start_date ? new Date(cmp.start_date) : null;
      const end = cmp.end_date ? new Date(cmp.end_date) : null;
      const endIn7 = end && (end - now) / 86400000 <= 7 && (end - now) >= 0;

      if (cmp.mounting_status !== 'Mounted' && start && start <= now) {
        alerts.push({
          id: cmp.id,
          site_code: cmp.site_code,
          client: cmp.client || 'Client',
          campaign: cmp.campaign_name || 'Campaign',
          tag: 'Mounting overdue',
          class: 'danger',
          link: '/campaigns'
        });
      }
      if (endIn7) {
        alerts.push({
          id: cmp.id,
          site_code: cmp.site_code,
          client: cmp.client || 'Client',
          campaign: cmp.campaign_name || 'Campaign',
          tag: 'Campaign ending soon',
          class: 'watch',
          link: '/campaigns'
        });
      }
      if (cmp.validation_15_date && new Date(cmp.validation_15_date) <= now) {
        alerts.push({
          id: cmp.id,
          site_code: cmp.site_code,
          client: cmp.client || 'Client',
          campaign: cmp.campaign_name || 'Campaign',
          tag: '15-Day validation due',
          class: 'danger',
          link: '/validations'
        });
      }
      if (cmp.final_validation_date && new Date(cmp.final_validation_date) <= now) {
        alerts.push({
          id: cmp.id,
          site_code: cmp.site_code,
          client: cmp.client || 'Client',
          campaign: cmp.campaign_name || 'Campaign',
          tag: 'Final validation due',
          class: 'danger',
          link: '/validations'
        });
      }
      if (cmp.invoice_status === 'Pending' || cmp.hard_copy_status === 'Pending') {
        alerts.push({
          id: cmp.id,
          site_code: cmp.site_code,
          client: cmp.client || 'Client',
          campaign: cmp.campaign_name || 'Campaign',
          tag: 'Invoice action required',
          class: 'watch',
          link: '/invoices'
        });
      }
    }

    // Real overdue electricity alerts
    for (const el of (electricityRows || [])) {
      if (el.due_date && new Date(el.due_date) < now) {
        alerts.push({
          id: el.id,
          site_code: el.site_code || 'Power',
          client: el.bill_type || 'Power Bill',
          campaign: `Due ₹${Number(el.amount || 0).toLocaleString('en-IN')}`,
          tag: 'Power bill overdue',
          class: 'danger',
          link: '/electricity'
        });
      }
    }

    res.json({
      kpis: {
        total_sites: totalSites,
        flagged_count: flaggedCount,
        active_sites: activeSites,
        nonactive_sites: nonActiveSites,
        active_campaigns: activeCampaigns,
        mounting_overdue: Number(cAll?.mounting_overdue || 0),
        validation_15_due: Number(cAll?.val15 || 0),
        final_validation_due: Number(cAll?.valFinal || 0),
        invoice_actions: Number(i?.pending || 0),
        open_proposals: Number(p?.open_proposals || 0),
        pending_bills: Number(e?.pending_bills || 0),
        electricity_overdue: Number(e?.overdue || 0),
        unpaid_electricity: Number(e?.unpaid || 0),
        average_occupancy: avgOccupancy,
        campaign_revenue: Number(cLive?.revenue || 0),
        gross_margin: Number(cLive?.margin || 0)
      },
      occupancy_buckets: occupancyBuckets,
      alerts: alerts.slice(0, 20)
    });
  } catch (err) {
    console.error('Dashboard error:', err.message);
    res.status(500).json({ message: 'Dashboard data error: ' + err.message });
  }
});
// Dedicated API Endpoints (must be registered BEFORE /api/:entity)

// Site Photos Upload & Delete
app.post('/api/sites/:id/images', auth, upload.array('files', 12), async (req, res) => {
  try {
    const site = (await q('SELECT * FROM sites WHERE id=?', [req.params.id]))[0];
    if (!site) return res.status(404).json({ message: 'Site not found' });
    let flags = {};
    try { flags = JSON.parse(site.flags || '{}') || {}; } catch {};
    if (Array.isArray(flags)) flags = { tags: flags };
    const existing = Array.isArray(flags.ppt_images) ? flags.ppt_images : [];
    const urls = (req.files || []).map(f => {
      const ext = path.extname(f.originalname || '');
      const final = f.path + ext;
      fs.renameSync(f.path, final);
      return `/uploads/${path.basename(final)}`;
    });
    flags.ppt_images = [...existing, ...urls];
    await q('UPDATE sites SET flags=?, updated_at=NOW() WHERE id=?', [JSON.stringify(flags), req.params.id]);
    res.json({ images: flags.ppt_images });
  } catch (err) {
    console.error('Image upload error:', err);
    res.status(500).json({ message: 'Upload error: ' + err.message });
  }
});

app.delete('/api/sites/:id/images/:index', auth, async (req, res) => {
  try {
    const site = (await q('SELECT * FROM sites WHERE id=?', [req.params.id]))[0];
    if (!site) return res.status(404).json({ message: 'Site not found' });
    let flags = {};
    try { flags = JSON.parse(site.flags || '{}') || {}; } catch {};
    if (Array.isArray(flags)) flags = { tags: flags };
    let existing = Array.isArray(flags.ppt_images) ? flags.ppt_images : [];
    const idx = Number(req.params.index);
    if (!isNaN(idx) && idx >= 0 && idx < existing.length) {
      existing.splice(idx, 1);
      flags.ppt_images = existing;
      await q('UPDATE sites SET flags=?, updated_at=NOW() WHERE id=?', [JSON.stringify(flags), req.params.id]);
    }
    res.json({ images: flags.ppt_images });
  } catch (err) {
    res.status(500).json({ message: 'Delete error: ' + err.message });
  }
});

// Fixed PPT Pages
app.get('/api/ppt-pages', auth, async (req, res) => {
  try {
    const rows = await q("SELECT setting_key, setting_value FROM settings WHERE setting_key IN ('ppt_first_page','ppt_second_last_page','ppt_last_page')");
    const m = Object.fromEntries(rows.map(r => [r.setting_key, r.setting_value]));
    res.json({
      first: m.ppt_first_page || '',
      second_last: m.ppt_second_last_page || '',
      last: m.ppt_last_page || ''
    });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

app.post('/api/ppt-pages/:key', auth, upload.single('file'), async (req, res) => {
  try {
    const map = { first: 'ppt_first_page', second_last: 'ppt_second_last_page', last: 'ppt_last_page' };
    const sk = map[req.params.key];
    if (!sk || !req.file) return res.status(400).json({ message: 'Invalid page upload' });
    const ext = path.extname(req.file.originalname || '');
    const final = req.file.path + ext;
    fs.renameSync(req.file.path, final);
    const url = `/uploads/${path.basename(final)}`;
    await q('INSERT INTO settings(setting_key,setting_value,updated_at) VALUES(?,?,NOW()) ON DUPLICATE KEY UPDATE setting_value=VALUES(setting_value),updated_at=NOW()', [sk, url]);
    res.json({ key: req.params.key, url });
  } catch (err) {
    res.status(500).json({ message: 'Page upload error: ' + err.message });
  }
});

app.delete('/api/ppt-pages/:key', auth, async (req, res) => {
  try {
    const map = { first: 'ppt_first_page', second_last: 'ppt_second_last_page', last: 'ppt_last_page' };
    const sk = map[req.params.key];
    if (sk) {
      await q('UPDATE settings SET setting_value="", updated_at=NOW() WHERE setting_key=?', [sk]);
    }
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

// User Management
app.get('/api/users', auth, async (req, res) => {
  try {
    const rows = await q('SELECT id, name, email, role, status, created_at, updated_at FROM users ORDER BY id ASC');
    res.json(rows);
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

app.post('/api/users', auth, admin, async (req, res) => {
  try {
    const { name, email, password, role = 'staff', status = 'active' } = req.body || {};
    if (!name || !email || !password) {
      return res.status(400).json({ message: 'Name, email, and password are required' });
    }
    const exists = await q('SELECT id FROM users WHERE email=? LIMIT 1', [email]);
    if (exists.length > 0) {
      return res.status(400).json({ message: 'A user with this email already exists' });
    }
    const hash = await bcrypt.hash(password, 10);
    const result = await q(
      'INSERT INTO users (name, email, password_hash, role, status, created_at, updated_at) VALUES (?, ?, ?, ?, ?, NOW(), NOW())',
      [name, email, hash, role, status]
    );
    res.json({ id: result.insertId, name, email, role, status });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

app.put('/api/users/:id', auth, admin, async (req, res) => {
  try {
    const { name, email, password, role, status } = req.body || {};
    const userId = req.params.id;
    if (password && String(password).trim()) {
      const hash = await bcrypt.hash(password, 10);
      await q('UPDATE users SET name=?, email=?, password_hash=?, role=?, status=?, updated_at=NOW() WHERE id=?', [name, email, hash, role, status, userId]);
    } else {
      await q('UPDATE users SET name=?, email=?, role=?, status=?, updated_at=NOW() WHERE id=?', [name, email, role, status, userId]);
    }
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

app.delete('/api/users/:id', auth, admin, async (req, res) => {
  try {
    const userId = req.params.id;
    if (String(req.user.id) === String(userId)) {
      return res.status(400).json({ message: 'You cannot delete your own account' });
    }
    await q('DELETE FROM users WHERE id=?', [userId]);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

// Settings
app.get('/api/settings', auth, async (req, res) => {
  try {
    const rows = await q('SELECT setting_key, setting_value FROM settings');
    res.json(Object.fromEntries(rows.map(r => [r.setting_key, r.setting_value])));
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

app.put('/api/settings', auth, admin, async (req, res) => {
  try {
    for (const [k, v] of Object.entries(req.body || {})) {
      await q('INSERT INTO settings(setting_key,setting_value,updated_at) VALUES(?,?,NOW()) ON DUPLICATE KEY UPDATE setting_value=VALUES(setting_value),updated_at=NOW()', [k, String(v ?? '')]);
    }
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

// Notifications
app.get('/api/notifications', auth, async (req, res) => {
  try {
    const rows = await q('SELECT * FROM notifications ORDER BY id DESC LIMIT 100');
    res.json(rows);
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

app.post('/api/notifications/read', auth, async (req, res) => {
  try {
    if (req.body?.id) await q('UPDATE notifications SET is_read=1 WHERE id=?', [req.body.id]);
    else await q('UPDATE notifications SET is_read=1');
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

app.post('/api/notifications/clear', auth, async (req, res) => {
  try {
    await q('DELETE FROM notifications');
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

app.delete('/api/notifications', auth, async (req, res) => {
  try {
    await q('DELETE FROM notifications');
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

// Activity Log
app.get('/api/activity', auth, async (req, res) => {
  try {
    const rows = await q('SELECT * FROM activity_log ORDER BY id DESC LIMIT 250');
    res.json(rows);
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

// Export & Import
app.get('/api/export/json', auth, async (req, res) => {
  try {
    const out = { format: 'site-control-webapp', exported_at: new Date().toISOString() };
    for (const e of Object.values(entities)) {
      out[e.table] = await q(`SELECT * FROM \`${e.table}\``);
    }
    res.setHeader('Content-Disposition', 'attachment; filename=site-control-backup.json');
    res.json(out);
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

function parseGps(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return { gps: '', latitude: null, longitude: null };
  const pair = raw.match(/^\s*(-?\d+(?:\.\d+)?)\s*[, ]\s*(-?\d+(?:\.\d+)?)\s*$/);
  if (pair) return { gps: `${parseFloat(pair[1])}, ${parseFloat(pair[2])}`, latitude: parseFloat(pair[1]), longitude: parseFloat(pair[2]) };
  const dms = /(\d+(?:\.\d+)?)\s*[°]\s*(\d+(?:\.\d+)?)?\s*['′]?\s*(\d+(?:\.\d+)?)?\s*["″]?\s*([NS]).*?(\d+(?:\.\d+)?)\s*[°]\s*(\d+(?:\.\d+)?)?\s*['′]?\s*(\d+(?:\.\d+)?)?\s*["″]?\s*([EW])/i.exec(raw);
  if (dms) {
    const cv = (d, m, s) => Number(d || 0) + Number(m || 0) / 60 + Number(s || 0) / 3600;
    const lat = cv(dms[1], dms[2], dms[3]) * (dms[4].toUpperCase() === 'S' ? -1 : 1);
    const lng = cv(dms[5], dms[6], dms[7]) * (dms[8].toUpperCase() === 'W' ? -1 : 1);
    return { gps: `${lat.toFixed(6)}, ${lng.toFixed(6)}`, latitude: lat, longitude: lng };
  }
  return { gps: raw, latitude: null, longitude: null };
}

function normalizeKey(str) {
  return String(str || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function getRowValue(row, keys, fallback = '') {
  const map = {};
  for (const k of Object.keys(row || {})) {
    map[normalizeKey(k)] = row[k];
  }
  for (const k of keys) {
    const nk = normalizeKey(k);
    if (map[nk] !== undefined && String(map[nk]).trim() !== '') {
      return map[nk];
    }
  }
  return fallback;
}

function parseAmount(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  const str = String(val || '').trim();
  if (!str) return 0;
  const lacMatch = str.match(/^([\d.]+)\s*(?:lac|lakh|l)s?$/i);
  if (lacMatch) {
    return Math.round(parseFloat(lacMatch[1]) * 100000);
  }
  const cleaned = str.replace(/[^0-9.]/g, '');
  return parseFloat(cleaned) || 0;
}

function parseDimensions(rawW, rawH, rawSize, rawSqFt) {
  let w = parseFloat(rawW) || null;
  let h = parseFloat(rawH) || null;
  const sizeStr = String(rawSize || '').trim();

  if ((!w || !h) && sizeStr) {
    const m = sizeStr.match(/(\d+(?:\.\d+)?)\s*(?:x|\*|X|by)\s*(\d+(?:\.\d+)?)/i);
    if (m) {
      if (!w) w = parseFloat(m[1]);
      if (!h) h = parseFloat(m[2]);
    }
  }

  let sqFt = parseFloat(rawSqFt) || null;
  if (!sqFt && w && h) {
    sqFt = Math.round(w * h * 100) / 100;
  }

  const finalSize = (w && h) ? `${w}x${h} ft` : (sizeStr || (sqFt ? `${sqFt} sq ft` : ''));
  return { w, h, sqFt, size: finalSize };
}

function parseAvailability(val) {
  if (val instanceof Date && !isNaN(val)) {
    const dd = String(val.getDate()).padStart(2, '0');
    const mm = String(val.getMonth() + 1).padStart(2, '0');
    return `${dd}.${mm}.${val.getFullYear()}`;
  }
  if (typeof val === 'number' && val > 20000) {
    const d = XLSX.SSF.parse_date_code(val);
    if (d) return `${String(d.d).padStart(2, '0')}.${String(d.m).padStart(2, '0')}.${d.y}`;
  }
  const str = String(val ?? '').trim().replace(/\.$/, '');
  if (!str) return 'Available';
  const isoMatch = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (isoMatch) {
    return `${isoMatch[3].padStart(2, '0')}.${isoMatch[2].padStart(2, '0')}.${isoMatch[1]}`;
  }
  const dmyMatch = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (dmyMatch) {
    return `${dmyMatch[1].padStart(2, '0')}.${dmyMatch[2].padStart(2, '0')}.${dmyMatch[3]}`;
  }
  const low = str.toLowerCase();
  if (low === 'yes' || low === 'y' || low === 'immediate' || low === 'ready' || low === 'vacant') return 'Available';
  if (low === 'no' || low === 'n' || low === 'occupied' || low === 'booked') return 'Occupied';
  return str;
}

function cleanStr(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function parseMysqlDate(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val)) {
    return val.toISOString().slice(0, 10);
  }
  if (typeof val === 'number' && val > 20000) {
    const d = XLSX.SSF.parse_date_code(val);
    if (d) {
      const yy = String(d.y).padStart(4, '0');
      const mm = String(d.m).padStart(2, '0');
      const dd = String(d.d).padStart(2, '0');
      return `${yy}-${mm}-${dd}`;
    }
  }
  const str = String(val).trim();
  if (!str) return null;
  const iso = str.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (iso) {
    return `${iso[1]}-${String(iso[2]).padStart(2, '0')}-${String(iso[3]).padStart(2, '0')}`;
  }
  const dmy = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (dmy) {
    return `${dmy[3]}-${String(dmy[2]).padStart(2, '0')}-${String(dmy[1]).padStart(2, '0')}`;
  }
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return parsed.toISOString().slice(0, 10);
  }
  return null;
}

app.post('/api/import/xlsx', auth, managerOrAdmin, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ message: 'No workbook provided' });
  try {
    const wb = XLSX.readFile(req.file.path, { cellDates: true });
    if (!wb.SheetNames || wb.SheetNames.length === 0) {
      return res.status(400).json({ message: 'Excel workbook contains no sheets' });
    }

    // Header keywords across languages and formats
    const headerKeywords = [
      'area', 'landmark', 'zone', 'locality',
      'location', 'address', 'sitename', 'site name',
      'media', 'mediatype', 'displaytype',
      'light', 'illumination',
      'w', 'width', 'h', 'height', 'size', 'dimension', 'sqft', 'sq ft',
      'rate', 'selling', 'adv fee', 'fee', 'price', 'amount', 'rent', 'cost',
      'avail', 'status',
      'sr no', 'srno', 'sno', 'serial',
      'site id', 'site code', 'siteid', 'sitecode', 'code',
      'latitude', 'longitude', 'gps', 'coords', 'coordinates'
    ];

    let chosenSheet = null;
    let bestHeaderIdx = 0;
    let highestScore = -1;

    // Scan sheets to find the one with the best header match
    for (const name of wb.SheetNames) {
      const sheet = wb.Sheets[name];
      if (!sheet || !sheet['!ref']) continue;
      const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
      for (let i = 0; i < Math.min(25, rawRows.length); i++) {
        const row = rawRows[i] || [];
        let score = 0;
        for (const cell of row) {
          const txt = String(cell || '').toLowerCase().replace(/[^a-z0-9]/g, '');
          if (!txt) continue;
          for (const kw of headerKeywords) {
            const cleanKw = kw.replace(/[^a-z0-9]/g, '');
            if (txt.includes(cleanKw)) {
              score++;
              break;
            }
          }
        }
        if (score > highestScore) {
          highestScore = score;
          chosenSheet = sheet;
          bestHeaderIdx = i;
        }
      }
    }

    if (!chosenSheet) {
      chosenSheet = wb.Sheets[wb.SheetNames[0]];
      bestHeaderIdx = 0;
    }

    // If a confident header row was found (>= 2 keywords), use it as the range origin
    const headerRange = highestScore >= 2 ? bestHeaderIdx : 0;
    const rows = XLSX.utils.sheet_to_json(chosenSheet, { range: headerRange, defval: '', raw: true });

    let updatedCount = 0, newCount = 0;

    // Load active sites for matching
    const existingSites = await q(`SELECT id, site_code, city, area, address, media_type, lighting, size, width, height,
      availability, monthly_rate, monthly_cost, latitude, longitude, gps, flags FROM sites WHERE record_status="active"`);

    // Determine current highest MB-XX site number
    let maxMbNum = 0;
    for (const st of existingSites) {
      const m = String(st.site_code || '').match(/MB[ -]?(\d+)/i);
      if (m) {
        const n = parseInt(m[1], 10);
        if (n > maxMbNum) maxMbNum = n;
      }
    }

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      let area = String(getRowValue(r, ['AREA', 'Area', 'Area / Landmark', 'Area/Landmark', 'Landmark', 'Zone', 'Locality', 'City/Area'])).trim();
      let location = String(getRowValue(r, ['LOCATION', 'Location', 'Full Address', 'Address', 'Site Location', 'Site Address', 'Location / Landmark', 'Road / Location', 'Site Name', 'Description'])).trim();

      // If one is missing, infer from the other
      if (!area && location) {
        const parts = location.split(/[-–,]/);
        if (parts.length > 1 && parts[0].trim().length > 2 && parts[0].trim().length < 35) {
          area = parts[0].trim();
        } else {
          area = 'Ahmedabad';
        }
      } else if (area && !location) {
        location = area;
      }

      if (!area && !location) continue; // Skip completely empty line

      // Media Type
      const rawMedia = String(getRowValue(r, ['MEDIA', 'Media', 'Media Type', 'MediaType', 'Type', 'Media_Type', 'Display Type'], '')).trim();
      let mediaType = 'Hoarding';
      if (/gantry/i.test(rawMedia)) mediaType = 'Gantry';
      else if (/unipole/i.test(rawMedia)) mediaType = 'Unipole';
      else if (/billboard/i.test(rawMedia)) mediaType = 'Billboard';
      else if (/dooh|digital|led/i.test(rawMedia)) mediaType = 'DOOH';
      else if (/kiosk/i.test(rawMedia)) mediaType = 'Kiosk';
      else if (/bridge/i.test(rawMedia)) mediaType = 'Bridge Panel';
      else if (rawMedia) mediaType = rawMedia;

      // Lighting
      const rawLight = String(getRowValue(r, ['LIGHT', 'Lighting', 'Light', 'Illumination', 'Light Type'], '')).trim();
      let lighting = 'BL';
      if (/front\s*lit|fl/i.test(rawLight)) lighting = 'FL';
      else if (/back\s*lit|bl/i.test(rawLight)) lighting = 'BL';
      else if (/non\s*lit|nl|unlit/i.test(rawLight)) lighting = 'NL';
      else if (/led|digital/i.test(rawLight)) lighting = 'LED';
      else if (rawLight) lighting = rawLight.toUpperCase();

      // Dimensions & Size
      const rawW = getRowValue(r, ['W', 'Width', 'W (ft)', 'Width (ft)', 'Width(ft)', 'W (FT)', 'Width in Feet'], '');
      const rawH = getRowValue(r, ['H', 'Height', 'H (ft)', 'Height (ft)', 'Height(ft)', 'H (FT)', 'Height in Feet'], '');
      const rawSize = getRowValue(r, ['Size', 'Dimensions', 'Size (ft)', 'Size (WxH)', 'Size(ft)', 'Dimension'], '');
      const rawSqFt = getRowValue(r, ['SQ FT', 'SQFT', 'Sq Ft', 'SqFt', 'Area (Sq Ft)', 'Area (Sqft)', 'Total Sqft', 'Total Sq Ft'], '');
      const dim = parseDimensions(rawW, rawH, rawSize, rawSqFt);

      // Availability
      const rawAvail = getRowValue(r, ['AVAILABLITY', 'AVAILABILITY', 'Availability', 'Status', 'Avail', 'Available From', 'Vacant From'], 'Available');
      const availability = parseAvailability(rawAvail);

      // Rate / Adv. Fee Per Month
      const rawSelling = getRowValue(r, [
        'Selling Amount', 'SellingAmount', 'Selling Amount (₹)', 'Adv. Fee Per Month', 'Adv Fee Per Month',
        'Adv Fee', 'Adv. Fee', 'Rate', 'Monthly Rate', 'Rate Per Month', 'Price', 'Rental', 'Cost', 'Amount'
      ], '0');
      const monthlyRate = parseAmount(rawSelling);

      // GPS & Coordinates
      const rawGps = getRowValue(r, ['Latitude Longitude', 'LatitudeLongitude', 'GPS', 'Coordinates', 'Latitude, Longitude', 'Coords', 'Lat/Long', 'Lat Long', 'Geo Coordinates'], '');
      const gpsObj = parseGps(rawGps);
      const lat = gpsObj.latitude || parseFloat(getRowValue(r, ['Latitude', 'Lat'], '')) || null;
      const lng = gpsObj.longitude || parseFloat(getRowValue(r, ['Longitude', 'Long', 'Lng', 'Lon'], '')) || null;
      const formattedGps = gpsObj.gps || (lat && lng ? `${lat}, ${lng}` : '');

      // Site Code in row
      const codeInRow = String(getRowValue(r, ['Site ID', 'site_code', 'SITE ID', 'Site Code', 'SiteCode', 'Code', 'ID'], '')).trim();

      // Match existing site
      let matched = null;

      // 1. Check direct site code match
      if (codeInRow) {
        const cleanCode = cleanStr(codeInRow);
        matched = existingSites.find(s => cleanStr(s.site_code) === cleanCode);
        if (!matched && /^\d+$/.test(codeInRow)) {
          const num = parseInt(codeInRow, 10);
          const mbVariant = `mb${String(num).padStart(2, '0')}`;
          matched = existingSites.find(s => cleanStr(s.site_code) === mbVariant || cleanStr(s.site_code) === `mb${num}`);
        }
      }

      // 2. Check exact or clean address match
      if (!matched && location) {
        const cleanLoc = cleanStr(location);
        matched = existingSites.find(s => cleanStr(s.address) === cleanLoc);
      }

      // 3. Check partial address containment
      if (!matched && location && cleanStr(location).length >= 10) {
        const cleanLoc = cleanStr(location);
        matched = existingSites.find(s => {
          const cleanAddr = cleanStr(s.address);
          if (!cleanAddr || cleanAddr.length < 10) return false;
          return cleanAddr.includes(cleanLoc) || cleanLoc.includes(cleanAddr);
        });
      }

      // 4. Check area + location landmark overlap
      if (!matched && area && location) {
        const cleanA = cleanStr(area);
        const cleanL = cleanStr(location);
        matched = existingSites.find(s => {
          if (cleanStr(s.area) !== cleanA) return false;
          const cleanAddr = cleanStr(s.address);
          return cleanAddr.includes(cleanL.slice(0, 15)) || cleanL.includes(cleanAddr.slice(0, 15));
        });
      }

      if (matched) {
        let oldFlags = {};
        try {
          oldFlags = typeof matched.flags === 'string' ? JSON.parse(matched.flags || '{}') : (matched.flags || {});
        } catch (e) { oldFlags = {}; }
        if (Array.isArray(oldFlags)) oldFlags = { ppt_images: oldFlags };

        const effectiveRate = monthlyRate || oldFlags.ppt_rate || matched.monthly_rate || 0;
        const mergedFlags = JSON.stringify({
          ...oldFlags,
          ppt_availability: availability,
          ppt_rate: effectiveRate,
          ppt_images: oldFlags.ppt_images || []
        });

        await q(`UPDATE sites SET
          area = COALESCE(NULLIF(?, ''), area),
          address = COALESCE(NULLIF(?, ''), address),
          media_type = COALESCE(NULLIF(?, ''), media_type),
          lighting = COALESCE(NULLIF(?, ''), lighting),
          size = COALESCE(NULLIF(?, ''), size),
          width = COALESCE(?, width),
          height = COALESCE(?, height),
          availability = ?,
          monthly_rate = ?,
          latitude = COALESCE(?, latitude),
          longitude = COALESCE(?, longitude),
          gps = COALESCE(NULLIF(?, ''), gps),
          flags = ?,
          updated_at = NOW()
          WHERE id = ?`, [
          area, location, mediaType, lighting, dim.size, dim.w, dim.h,
          availability, effectiveRate, lat, lng, formattedGps, mergedFlags, matched.id
        ]);
        updatedCount++;
      } else {
        // Generate new site code
        let newSiteCode = codeInRow && !/^\d+$/.test(codeInRow) ? codeInRow : null;
        if (!newSiteCode) {
          maxMbNum++;
          newSiteCode = `MB-${String(maxMbNum).padStart(2, '0')}`;
        }

        const flags = JSON.stringify({
          ppt_availability: availability,
          ppt_rate: monthlyRate,
          ppt_images: []
        });

        await q(`INSERT INTO sites (
          site_code, city, area, address, media_type, lighting, size, width, height,
          availability, monthly_cost, monthly_rate, latitude, longitude, gps, flags, record_status, created_at, updated_at
        ) VALUES (?, 'Ahmedabad', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'active', NOW(), NOW())`, [
          newSiteCode, area, location, mediaType, lighting, dim.size, dim.w, dim.h,
          availability, monthlyRate, monthlyRate, lat, lng, formattedGps, flags
        ]);
        newCount++;
      }
    }

    res.json({
      success: true,
      updated: updatedCount,
      created: newCount,
      rows: updatedCount + newCount,
      message: `Successfully processed ${updatedCount + newCount} sites (${updatedCount} updated, ${newCount} created).`
    });
  } catch (err) {
    console.error('Import XLSX error:', err);
    res.status(500).json({ message: 'Excel import error: ' + err.message });
  }
});

app.post('/api/import/electricity-xlsx', auth, managerOrAdmin, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ message: 'No workbook provided' });
  try {
    const wb = XLSX.readFile(req.file.path, { cellDates: true });
    if (!wb.SheetNames || wb.SheetNames.length === 0) {
      return res.status(400).json({ message: 'Excel workbook contains no sheets' });
    }

    const headerKeywords = [
      'meter', 'service', 'tnumber', 't_number', 'tno',
      'billingmonth', 'billdate', 'duedate', 'paiddate',
      'units', 'provider', 'billtype', 'paymentamount', 'billamount',
      'amount', 'paymentstatus', 'paymentref', 'status',
      'sitecode', 'siteid', 'location', 'size', 'ssv', 'mb'
    ];

    const candidateSheets = [];
    for (const name of wb.SheetNames) {
      const sheet = wb.Sheets[name];
      if (!sheet || !sheet['!ref']) continue;
      const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
      let bestScore = 0, bestIdx = 0;
      for (let i = 0; i < Math.min(25, rawRows.length); i++) {
        const row = rawRows[i] || [];
        let score = 0;
        for (const cell of row) {
          const txt = String(cell || '').toLowerCase().replace(/[^a-z0-9]/g, '');
          if (!txt) continue;
          for (const kw of headerKeywords) {
            if (txt.includes(kw)) {
              score++;
              break;
            }
          }
        }
        if (score > bestScore) {
          bestScore = score;
          bestIdx = i;
        }
      }
      if (bestScore >= 2) {
        candidateSheets.push({ name, sheet, headerIdx: bestIdx, score: bestScore });
      }
    }

    if (candidateSheets.length === 0) {
      candidateSheets.push({ name: wb.SheetNames[0], sheet: wb.Sheets[wb.SheetNames[0]], headerIdx: 0, score: 0 });
    }

    let updatedCount = 0, newCount = 0;

    // Load active sites and active electricity records
    const existingSites = await q(`SELECT id, site_code, area, address, size, width, height, meter_no FROM sites WHERE record_status="active"`);
    const existingBills = await q(`SELECT id, site_id, site_code, meter_no, service_number, t_number, billing_month, due_date, amount, payment_status, location FROM electricity WHERE record_status="active"`);

    for (const cs of candidateSheets) {
      const sheetName = cs.name;
      const sheetHint = /ssv/i.test(sheetName) ? 'SSV' : (/mb\b|media\s*buzz/i.test(sheetName) ? 'MB' : '');
      const rows = XLSX.utils.sheet_to_json(cs.sheet, { range: cs.headerIdx, defval: '', raw: true });

      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];

        const siteCode = String(getRowValue(r, ['SITE CODE', 'Site Code', 'SiteCode', 'Site ID', 'SiteID', 'Code', 'Site'], '')).trim();
        const location = String(getRowValue(r, ['LOCATION', 'Location', 'Address', 'Full Address', 'Site Location', 'Landmark'], '')).trim();
        const size = String(getRowValue(r, ['SIZE', 'Size', 'Dimensions'], '')).trim();
        const meterNo = String(getRowValue(r, ['METER NO', 'Meter No', 'Meter Number', 'Meter #', 'Meter', 'MTR'], '')).trim();
        const serviceNo = String(getRowValue(r, ['SERVICE NUMBER', 'Service Number', 'Service No', 'Consumer No', 'Consumer Number', 'Customer ID', 'Account Number'], '')).trim();
        const tNumber = String(getRowValue(r, ['T NUMBER', 'T Number', 'T-No', 'TNo', 'Tariff Number'], '')).trim();

        // 1. Check explicit provider/bill header variations
        const rawBillType = String(getRowValue(r, [
          'BILL / PROVIDER', 'Bill / Provider', 'Bill/Provider', 'Bill/provider', 'BILL PROVIDER', 'Bill Provider', 'bill provider',
          'SSV / MB', 'SSV/MB', 'SSV MB', 'SSV', 'MB', 'SSV/MB BILL', 'SSV / MB BILL',
          'BILL TYPE', 'Bill Type', 'bill type', 'BillType', 'Bill_Type',
          'PROVIDER', 'Provider', 'provider', 'Discom', 'Company', 'Firm', 'Entity', 'Party', 'Biller',
          'BILL', 'Bill', 'bill'
        ], '')).trim();

        let detectedBillType = rawBillType;

        // 2. If not found, inspect columns that might contain provider or firm info
        if (!detectedBillType) {
          for (const k of Object.keys(r || {})) {
            const nk = normalizeKey(k);
            if (nk.includes('provider') || nk.includes('discom') || nk.includes('biller') || nk.includes('ssv') || nk.includes('billtype') || nk.includes('firm') || nk.includes('entity') || nk.includes('party')) {
              const val = String(r[k] || '').trim();
              if (val) {
                detectedBillType = val;
                break;
              }
            }
          }
        }

        // 3. If still not found, check if any cell value strictly equals 'SSV' or 'MB'
        if (!detectedBillType) {
          for (const k of Object.keys(r || {})) {
            const val = String(r[k] || '').trim().toUpperCase();
            if (val === 'SSV' || val === 'MB') {
              detectedBillType = val;
              break;
            }
          }
        }

        // 4. Normalize provider value (SSV and MB as primary)
        let billType = sheetHint || 'SSV';
        if (detectedBillType) {
          const cleanUpper = detectedBillType.toUpperCase();
          if (/SSV/i.test(cleanUpper)) {
            billType = 'SSV';
          } else if (/^MB\b/i.test(cleanUpper) || /MEDIA\s*BUZZ/i.test(cleanUpper) || cleanUpper === 'MB') {
            billType = 'MB';
          } else if (/TORRENT/i.test(cleanUpper)) {
            billType = 'Torrent Power';
          } else if (/PGVCL/i.test(cleanUpper)) {
            billType = 'PGVCL';
          } else if (/UGVCL/i.test(cleanUpper)) {
            billType = 'UGVCL';
          } else if (/MGVCL/i.test(cleanUpper)) {
            billType = 'MGVCL';
          } else if (/DGVCL/i.test(cleanUpper)) {
            billType = 'DGVCL';
          } else {
            billType = detectedBillType;
          }
        }

        let billingMonth = String(getRowValue(r, ['BILLING MONTH', 'Billing Month', 'Month', 'Bill Month', 'Period'], '')).trim();
        const rawBillDate = getRowValue(r, ['BILL DATE', 'Bill Date', 'Date of Bill', 'Issue Date'], '');
        const billDate = parseMysqlDate(rawBillDate);
        const rawDueDate = getRowValue(r, ['DUE DATE', 'Due Date', 'Payment Due Date', 'Last Date'], '');
        let dueDate = parseMysqlDate(rawDueDate);

        // Fallback due date if missing (due_date is NOT NULL in MySQL)
        if (!dueDate) {
          if (billDate) {
            const bd = new Date(billDate);
            bd.setDate(bd.getDate() + 15);
            dueDate = bd.toISOString().slice(0, 10);
          } else {
            const now = new Date();
            now.setDate(now.getDate() + 15);
            dueDate = now.toISOString().slice(0, 10);
          }
        }

        // Fallback billing month if missing
        if (!billingMonth) {
          if (billDate) {
            billingMonth = new Date(billDate).toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
          } else if (dueDate) {
            billingMonth = new Date(dueDate).toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
          } else {
            billingMonth = new Date().toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
          }
        }

        const units = parseFloat(getRowValue(r, ['UNITS', 'Units', 'Consumption', 'KWH', 'Total Units'], '0')) || 0;
        const rate = parseFloat(getRowValue(r, ['RATE', 'Rate', 'Unit Rate', 'Tariff Rate'], '0')) || 0;
        const otherCharges = parseFloat(getRowValue(r, ['OTHER CHARGES', 'Other Charges', 'Fixed Charges', 'Surcharge'], '0')) || 0;

        const rawAmount = getRowValue(r, ['PAYMENT AMOUNT (₹)', 'PAYMENT AMOUNT', 'Payment Amount', 'AMOUNT', 'Amount', 'Bill Amount', 'Total Amount', 'Net Amount', 'Payable Amount'], '0');
        let amount = parseAmount(rawAmount);
        if (amount === 0 && units > 0 && rate > 0) {
          amount = Math.round((units * rate + otherCharges) * 100) / 100;
        }

        const rawStatus = String(getRowValue(r, ['STATUS', 'Status', 'Payment Status', 'Paid Status'], '')).trim();
        const rawPaidDate = getRowValue(r, ['PAID DATE', 'Paid Date', 'Payment Date', 'Cleared Date'], '');
        const paidDate = parseMysqlDate(rawPaidDate);
        const paymentRef = String(getRowValue(r, ['PAYMENT REF', 'Payment Ref', 'Payment Reference', 'UTR', 'Ref No', 'Transaction ID', 'Cheque No'], '')).trim();
        const notes = String(getRowValue(r, ['NOTES', 'Notes', 'Remarks', 'Comments', 'Description'], '')).trim();

        let paymentStatus = 'Pending';
        if (/paid|cleared|done|yes/i.test(rawStatus)) paymentStatus = 'Paid';
        else if (/overdue/i.test(rawStatus)) paymentStatus = 'Overdue';
        else if (paidDate || paymentRef) paymentStatus = 'Paid';
        else if (dueDate && new Date(dueDate) < new Date()) paymentStatus = 'Pending';

        // Skip row if completely empty
        if (!siteCode && !location && !meterNo && !serviceNo && amount === 0) continue;

        // Link to sites table
        let matchedSite = null;
        if (siteCode) {
          const cleanC = cleanStr(siteCode);
          matchedSite = existingSites.find(s => cleanStr(s.site_code) === cleanC);
          if (!matchedSite && /^\d+$/.test(siteCode)) {
            const num = parseInt(siteCode, 10);
            matchedSite = existingSites.find(s => cleanStr(s.site_code) === `mb${String(num).padStart(2, '0')}` || cleanStr(s.site_code) === `mb${num}`);
          }
        }
        if (!matchedSite && meterNo) {
          matchedSite = existingSites.find(s => cleanStr(s.meter_no) === cleanStr(meterNo));
        }
        if (!matchedSite && location) {
          matchedSite = existingSites.find(s => cleanStr(s.address) === cleanStr(location));
        }

        const finalSiteId = matchedSite?.id || null;
        const finalSiteCode = siteCode || matchedSite?.site_code || '';
        const finalLocation = location || matchedSite?.address || matchedSite?.area || '';
        const finalSize = size || matchedSite?.size || (matchedSite?.width && matchedSite?.height ? `${matchedSite.width}x${matchedSite.height} ft` : '');
        const finalMeterNo = meterNo || matchedSite?.meter_no || '';
        const finalServiceNo = serviceNo || '';
        const finalTNumber = tNumber || '';

        // Check if site meter_no can be backfilled
        if (matchedSite && meterNo && !matchedSite.meter_no) {
          await q('UPDATE sites SET meter_no=? WHERE id=?', [meterNo, matchedSite.id]);
        }

        // Check if bill already exists - match by unique service_number, t_number, meter_no, site_code or location
        let matchedBill = null;
        const cleanMonth = cleanStr(billingMonth);

        if (finalServiceNo) {
          const cleanSrv = cleanStr(finalServiceNo);
          matchedBill = existingBills.find(b => cleanStr(b.service_number) === cleanSrv && (!cleanMonth || cleanStr(b.billing_month) === cleanMonth));
          if (!matchedBill) {
            matchedBill = existingBills.find(b => cleanStr(b.service_number) === cleanSrv);
          }
        }
        if (!matchedBill && finalTNumber) {
          const cleanT = cleanStr(finalTNumber);
          matchedBill = existingBills.find(b => cleanStr(b.t_number) === cleanT && (!cleanMonth || cleanStr(b.billing_month) === cleanMonth));
          if (!matchedBill) {
            matchedBill = existingBills.find(b => cleanStr(b.t_number) === cleanT);
          }
        }
        if (!matchedBill && finalMeterNo) {
          const cleanMtr = cleanStr(finalMeterNo);
          matchedBill = existingBills.find(b => cleanStr(b.meter_no) === cleanMtr && (!cleanMonth || cleanStr(b.billing_month) === cleanMonth));
          if (!matchedBill) {
            matchedBill = existingBills.find(b => cleanStr(b.meter_no) === cleanMtr);
          }
        }
        if (!matchedBill && finalSiteCode && cleanMonth) {
          matchedBill = existingBills.find(b => cleanStr(b.site_code) === cleanStr(finalSiteCode) && cleanStr(b.billing_month) === cleanMonth);
        }
        if (!matchedBill && finalLocation) {
          const cleanLoc = cleanStr(finalLocation);
          matchedBill = existingBills.find(b => cleanStr(b.location) === cleanLoc);
        }

      if (matchedBill) {
        await q(`UPDATE electricity SET
          site_id = COALESCE(?, site_id),
          site_code = COALESCE(NULLIF(?, ''), site_code),
          location = COALESCE(NULLIF(?, ''), location),
          meter_no = COALESCE(NULLIF(?, ''), meter_no),
          size = COALESCE(NULLIF(?, ''), size),
          service_number = COALESCE(NULLIF(?, ''), service_number),
          t_number = COALESCE(NULLIF(?, ''), t_number),
          bill_type = COALESCE(NULLIF(?, ''), bill_type),
          payment_amount = ?,
          amount = ?,
          billing_month = ?,
          bill_date = COALESCE(?, bill_date),
          due_date = ?,
          units = ?,
          rate = ?,
          other_charges = ?,
          payment_status = ?,
          paid_date = COALESCE(?, paid_date),
          payment_reference = COALESCE(NULLIF(?, ''), payment_reference),
          notes = COALESCE(NULLIF(?, ''), notes),
          record_status = 'active',
          updated_at = NOW()
          WHERE id = ?`, [
          finalSiteId, finalSiteCode, finalLocation, finalMeterNo, finalSize, finalServiceNo, finalTNumber,
          billType, amount, amount, billingMonth, billDate, dueDate, units, rate, otherCharges,
          paymentStatus, paidDate, paymentRef, notes, matchedBill.id
        ]);
        updatedCount++;
      } else {
        await q(`INSERT INTO electricity (
          site_id, site_code, location, meter_no, size, service_number, t_number,
          bill_type, payment_amount, billing_month, bill_date, due_date, units, rate, other_charges,
          amount, payment_status, paid_date, payment_reference, notes, record_status, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'active', NOW(), NOW())`, [
          finalSiteId, finalSiteCode, finalLocation, finalMeterNo, finalSize, finalServiceNo, finalTNumber,
          billType, amount, billingMonth, billDate, dueDate, units, rate, otherCharges,
          amount, paymentStatus, paidDate, paymentRef, notes
        ]);
        newCount++;
      }
    }
  }

  res.json({
      success: true,
      updated: updatedCount,
      created: newCount,
      rows: updatedCount + newCount,
      message: `Successfully processed ${updatedCount + newCount} electricity bills (${updatedCount} updated, ${newCount} created).`
    });
  } catch (err) {
    console.error('Import Electricity XLSX error:', err);
    res.status(500).json({ message: 'Electricity bill import error: ' + err.message });
  }
});

app.post('/api/import/campaigns-xlsx', auth, managerOrAdmin, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ message: 'No workbook provided' });
  try {
    const wb = XLSX.readFile(req.file.path, { cellDates: true });
    if (!wb.SheetNames || wb.SheetNames.length === 0) {
      return res.status(400).json({ message: 'Excel workbook contains no sheets' });
    }

    const headerKeywords = [
      'site', 'code', 'month', 'date', 'client', 'agency', 'display', 'vendor',
      'location', 'size', 'type', 'start', 'end', 'days', 'advt',
      'fees', 'printing', 'mounting', 'total', 'amount', 'po', 'bill', 'pending'
    ];

    const candidateSheets = [];
    for (const name of wb.SheetNames) {
      const sheet = wb.Sheets[name];
      if (!sheet || !sheet['!ref']) continue;
      const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
      let bestScore = 0, bestIdx = 0;
      for (let i = 0; i < Math.min(25, rawRows.length); i++) {
        const row = rawRows[i] || [];
        let score = 0;
        for (const cell of row) {
          const txt = String(cell || '').toLowerCase().replace(/[^a-z0-9]/g, '');
          if (!txt) continue;
          for (const kw of headerKeywords) {
            if (txt.includes(kw)) {
              score++;
              break;
            }
          }
        }
        if (score > bestScore) {
          bestScore = score;
          bestIdx = i;
        }
      }
      if (bestScore >= 2) {
        candidateSheets.push({ name, sheet, headerIdx: bestIdx, score: bestScore });
      }
    }

    if (candidateSheets.length === 0) {
      candidateSheets.push({ name: wb.SheetNames[0], sheet: wb.Sheets[wb.SheetNames[0]], headerIdx: 0, score: 0 });
    }

    await ensureCampaignColumns();
    const campCols = await tableColumns('campaigns');

    let updatedCount = 0, newCount = 0;
    const existingSites = await q('SELECT id, site_code, address, area, width, height, size FROM sites WHERE record_status="active"');

    const selectCols = ['id', 'site_code', 'client', 'start_date', 'end_date'];
    ['month', 'booking_date', 'display', 'vendor_name', 'location', 'po'].forEach(c => {
      if (campCols.has(c)) selectCols.push(c);
    });
    const existingCampaigns = await q(`SELECT ${selectCols.join(', ')} FROM campaigns WHERE record_status="active"`);

    for (const cs of candidateSheets) {
      const rows = XLSX.utils.sheet_to_json(cs.sheet, { range: cs.headerIdx, defval: '', raw: true });

      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const getVal = (patterns) => {
          for (const p of patterns) {
            for (const key of Object.keys(r)) {
              const cleanKey = key.trim().toLowerCase().replace(/[^a-z0-9]/g, '');
              const cleanPattern = p.toLowerCase().replace(/[^a-z0-9]/g, '');
              if (cleanKey === cleanPattern || cleanKey.includes(cleanPattern)) {
                const val = r[key];
                if (val !== undefined && val !== null && String(val).trim() !== '') {
                  return val;
                }
              }
            }
          }
          return '';
        };

        const parseDate = (val) => {
          if (!val) return null;
          if (val instanceof Date) {
            if (isNaN(val.getTime())) return null;
            return val.toISOString().slice(0, 10);
          }
          if (typeof val === 'number') {
            const parsed = XLSX.SSF.parse_date_code(val);
            if (parsed) {
              const m = String(parsed.m).padStart(2, '0');
              const d = String(parsed.d).padStart(2, '0');
              return `${parsed.y}-${m}-${d}`;
            }
          }
          const s = String(val).trim();
          const dmy = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
          if (dmy) {
            const day = dmy[1].padStart(2, '0');
            const month = dmy[2].padStart(2, '0');
            let year = dmy[3];
            if (year.length === 2) year = '20' + year;
            return `${year}-${month}-${day}`;
          }
          const dt = new Date(s);
          if (!isNaN(dt.getTime())) {
            return dt.toISOString().slice(0, 10);
          }
          return null;
        };

        const parseNum = (val) => {
          if (typeof val === 'number') return isNaN(val) ? 0 : val;
          const cleaned = String(val || '').replace(/₹|,|\/|-/g, '').trim();
          const n = parseFloat(cleaned);
          return isNaN(n) ? 0 : n;
        };

        const siteCode = String(getVal(['site code', 'sitecode', 'site_code', 'site no', 'siteno', 'code'])).trim();
        const month = String(getVal(['month', 'billing_month', 'mon'])).trim();
        const dateVal = parseDate(getVal(['date', 'booking_date', 'booking date', 'dt']));
        const client = String(getVal(['client/agency name', 'client / agency name', 'client/agency', 'client_name', 'client', 'agency name', 'agency'])).trim();
        const display = String(getVal(['display', 'brand', 'campaign_name', 'campaign name', 'creative'])).trim();
        const vendorName = String(getVal(['vendor name', 'vendor_name', 'vendor'])).trim();
        const location = String(getVal(['location', 'site location', 'address', 'area', 'landmark'])).trim();
        
        let width = parseNum(getVal(['w', 'width']));
        let height = parseNum(getVal(['h', 'height']));
        let size = String(getVal(['size', 'dimension'])).trim();
        if (!size && width > 0 && height > 0) {
          size = `${width}x${height} ft`;
        } else if (size && (!width || !height)) {
          const dimMatch = size.match(/(\d+(?:\.\d+)?)\s*[xX*]\s*(\d+(?:\.\d+)?)/);
          if (dimMatch) {
            if (!width) width = parseFloat(dimMatch[1]);
            if (!height) height = parseFloat(dimMatch[2]);
          }
        }

        const type = String(getVal(['type', 'media type', 'media_type', 'media']) || 'Hoarding').trim();
        const startDate = parseDate(getVal(['start date', 'start_date', 'from date', 'from', 'start']));
        const endDate = parseDate(getVal(['end date', 'end_date', 'to date', 'to', 'end']));
        
        let days = parseInt(getVal(['days', 'duration_days', 'total days', 'duration']), 10);
        if (isNaN(days) || days <= 0) {
          if (startDate && endDate) {
            const diffMs = new Date(endDate).getTime() - new Date(startDate).getTime();
            days = Math.max(1, Math.round(diffMs / (1000 * 60 * 60 * 24)) + 1);
          } else {
            days = 30;
          }
        }

        const advtFees = parseNum(getVal(['advt. fees per month', 'advt fees per month', 'advt fees', 'advt. fees', 'advt_fees', 'fees', 'rate']));
        const printingMounting = parseNum(getVal(['printing & mounting', 'printing and mounting', 'printing & mounting cost', 'printing_mounting_cost', 'p&m', 'printing cost', 'mounting cost']));
        let totalAmount = parseNum(getVal(['total amount', 'total_amount', 'total', 'amount', 'revenue']));
        if (totalAmount === 0 && (advtFees > 0 || printingMounting > 0)) {
          totalAmount = advtFees + printingMounting;
        }

        const po = String(getVal(['po', 'po no', 'po number', 'po_number', 'purchase order'])).trim();
        const bill = String(getVal(['bill', 'bill no', 'bill_no', 'invoice no', 'invoice_no', 'bill status'])).trim();
        const pending = parseNum(getVal(['pending', 'pending amount', 'pending_amount', 'balance']));

        // Skip completely empty rows
        if (!siteCode && !client && !location && !display && totalAmount === 0 && !startDate) {
          continue;
        }

        // Link to sites table
        let matchedSite = null;
        if (siteCode) {
          const cleanC = cleanStr(siteCode);
          matchedSite = existingSites.find(s => cleanStr(s.site_code) === cleanC);
          if (!matchedSite && /^\d+$/.test(siteCode)) {
            const num = parseInt(siteCode, 10);
            matchedSite = existingSites.find(s => cleanStr(s.site_code) === `mb${String(num).padStart(2, '0')}` || cleanStr(s.site_code) === `mb${num}`);
          }
        }
        const finalSiteId = matchedSite?.id || null;
        const finalSiteCode = siteCode || matchedSite?.site_code || '';

        // Match existing campaign - take latest entry per site and do not create repeated entries
        const matched = existingCampaigns.find(c => {
          if (po && c.po && String(c.po).toLowerCase() === po.toLowerCase()) return true;
          if (finalSiteCode && c.site_code && String(c.site_code).toLowerCase() === finalSiteCode.toLowerCase()) return true;
          if (client && display && location &&
              String(c.client || '').toLowerCase() === client.toLowerCase() &&
              String(c.display || '').toLowerCase() === display.toLowerCase() &&
              String(c.location || '').toLowerCase() === location.toLowerCase()) return true;
          return false;
        });

        if (matched) {
          const updateFields = [
            'site_id = COALESCE(?, site_id)',
            'site_code = COALESCE(NULLIF(?, ""), site_code)'
          ];
          const updateVals = [finalSiteId, finalSiteCode];

          if (campCols.has('month')) { updateFields.push('month = COALESCE(NULLIF(?, ""), month)'); updateVals.push(month); }
          if (campCols.has('booking_date')) { updateFields.push('booking_date = COALESCE(?, booking_date)'); updateVals.push(dateVal); }
          if (campCols.has('client')) { updateFields.push('client = COALESCE(NULLIF(?, ""), client)'); updateVals.push(client); }
          if (campCols.has('display')) { updateFields.push('display = COALESCE(NULLIF(?, ""), display)'); updateVals.push(display); }
          if (campCols.has('campaign_name')) { updateFields.push('campaign_name = COALESCE(NULLIF(?, ""), campaign_name)'); updateVals.push(display); }
          if (campCols.has('brand')) { updateFields.push('brand = COALESCE(NULLIF(?, ""), brand)'); updateVals.push(client); }
          if (campCols.has('vendor_name')) { updateFields.push('vendor_name = COALESCE(NULLIF(?, ""), vendor_name)'); updateVals.push(vendorName); }
          if (campCols.has('location')) { updateFields.push('location = COALESCE(NULLIF(?, ""), location)'); updateVals.push(location); }
          if (campCols.has('width')) { updateFields.push('width = COALESCE(?, width)'); updateVals.push(width || null); }
          if (campCols.has('height')) { updateFields.push('height = COALESCE(?, height)'); updateVals.push(height || null); }
          if (campCols.has('size')) { updateFields.push('size = COALESCE(NULLIF(?, ""), size)'); updateVals.push(size); }
          if (campCols.has('type')) { updateFields.push('type = COALESCE(NULLIF(?, ""), type)'); updateVals.push(type); }
          if (campCols.has('start_date')) { updateFields.push('start_date = COALESCE(?, start_date)'); updateVals.push(startDate); }
          if (campCols.has('end_date')) { updateFields.push('end_date = COALESCE(?, end_date)'); updateVals.push(endDate); }
          if (campCols.has('days')) { updateFields.push('days = ?'); updateVals.push(days); }
          if (campCols.has('advt_fees')) { updateFields.push('advt_fees = ?'); updateVals.push(advtFees); }
          if (campCols.has('printing_mounting_cost')) { updateFields.push('printing_mounting_cost = ?'); updateVals.push(printingMounting); }
          if (campCols.has('total_amount')) { updateFields.push('total_amount = ?'); updateVals.push(totalAmount); }
          if (campCols.has('revenue')) { updateFields.push('revenue = ?'); updateVals.push(totalAmount); }
          if (campCols.has('po')) { updateFields.push('po = COALESCE(NULLIF(?, ""), po)'); updateVals.push(po); }
          if (campCols.has('bill')) { updateFields.push('bill = COALESCE(NULLIF(?, ""), bill)'); updateVals.push(bill); }
          if (campCols.has('pending')) { updateFields.push('pending = ?'); updateVals.push(pending); }
          updateFields.push('record_status = "active"', 'updated_at = NOW()');
          updateVals.push(matched.id);

          await q(`UPDATE campaigns SET ${updateFields.join(', ')} WHERE id = ?`, updateVals);
          matched.site_code = finalSiteCode;
          matched.start_date = startDate;
          matched.end_date = endDate;
          matched.month = month;
          updatedCount++;
        } else {
          const bookingCode = `MB-BK-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`;
          const insertData = {
            booking_code: bookingCode,
            site_id: finalSiteId,
            site_code: finalSiteCode,
            record_status: 'active'
          };
          if (campCols.has('month')) insertData.month = month;
          if (campCols.has('booking_date')) insertData.booking_date = dateVal;
          if (campCols.has('client')) insertData.client = client;
          if (campCols.has('display')) insertData.display = display;
          if (campCols.has('campaign_name')) insertData.campaign_name = display;
          if (campCols.has('brand')) insertData.brand = client;
          if (campCols.has('vendor_name')) insertData.vendor_name = vendorName;
          if (campCols.has('location')) insertData.location = location;
          if (campCols.has('width')) insertData.width = width || null;
          if (campCols.has('height')) insertData.height = height || null;
          if (campCols.has('size')) insertData.size = size;
          if (campCols.has('type')) insertData.type = type;
          if (campCols.has('start_date')) insertData.start_date = startDate;
          if (campCols.has('end_date')) insertData.end_date = endDate;
          if (campCols.has('days')) insertData.days = days;
          if (campCols.has('advt_fees')) insertData.advt_fees = advtFees;
          if (campCols.has('printing_mounting_cost')) insertData.printing_mounting_cost = printingMounting;
          if (campCols.has('total_amount')) insertData.total_amount = totalAmount;
          if (campCols.has('revenue')) insertData.revenue = totalAmount;
          if (campCols.has('po')) insertData.po = po;
          if (campCols.has('bill')) insertData.bill = bill;
          if (campCols.has('pending')) insertData.pending = pending;

          const cols = Object.keys(insertData);
          const placeholders = cols.map(() => '?').join(', ');
          const vals = cols.map(c => insertData[c]);
          const insertRes = await q(
            `INSERT INTO campaigns (${cols.map(c => '`' + c + '`').join(', ')}, created_at, updated_at) VALUES (${placeholders}, NOW(), NOW())`,
            vals
          );
          existingCampaigns.push({
            id: insertRes?.insertId,
            site_code: finalSiteCode,
            start_date: startDate,
            end_date: endDate,
            month,
            po, client, display, location
          });
          newCount++;
        }
      }
    }

    // Auto-sync all linked campaigns and site availability after Excel campaign import
    await syncAllLinkedCampaigns(q);
    await syncSiteAvailability();

    res.json({
      success: true,
      updated: updatedCount,
      created: newCount,
      rows: updatedCount + newCount,
      message: `Successfully processed ${updatedCount + newCount} campaigns (${updatedCount} updated, ${newCount} created). Combined & single linked sites auto-booked.`
    });
  } catch (err) {
    console.error('Import Campaigns XLSX error:', err);
    res.status(500).json({ message: 'Campaign import error: ' + err.message });
  }
});

app.post('/api/import/json', auth, managerOrAdmin, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ message: 'No JSON file provided' });
  try {
    const raw = JSON.parse(fs.readFileSync(req.file.path, 'utf8'));
    let totalImported = 0;
    
    // If it's a backup with table keys
    if (raw.sites && Array.isArray(raw.sites)) {
      for (const s of raw.sites) {
        if (!s.site_code) continue;
        const flags = typeof s.flags === 'string' ? s.flags : JSON.stringify(s.flags || {});
        await q(`INSERT INTO sites (
          site_code, city, area, address, size, media_type, lighting, facing,
          ownership, availability, vendor_name, meter_no, monthly_cost, monthly_rate,
          latitude, longitude, gps, notes, flags, record_status, created_at, updated_at
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?, 'active', NOW(), NOW())
        ON DUPLICATE KEY UPDATE city=VALUES(city), area=VALUES(area), address=VALUES(address),
        size=VALUES(size), media_type=VALUES(media_type), lighting=VALUES(lighting),
        availability=VALUES(availability), monthly_rate=VALUES(monthly_rate),
        latitude=VALUES(latitude), longitude=VALUES(longitude), flags=VALUES(flags)`, [
          s.site_code, s.city || 'Ahmedabad', s.area || '', s.address || '', s.size || '',
          s.media_type || 'Hoarding', s.lighting || 'BL', s.facing || '', s.ownership || 'Owned',
          s.availability || 'Available', s.vendor_name || '', s.meter_no || '',
          Number(s.monthly_cost || 0), Number(s.monthly_rate || 0), s.latitude || null, s.longitude || null,
          s.gps || '', s.notes || '', flags
        ]);
        totalImported++;
      }
    }
    
    // Also restore clients, campaigns, electricity, vendors if present
    for (const ent of ['clients', 'campaigns', 'electricity', 'vendors', 'invoices']) {
      if (Array.isArray(raw[ent])) {
        for (const item of raw[ent]) {
          const keys = Object.keys(item).filter(k => k !== 'id');
          if (!keys.length) continue;
          const vals = keys.map(k => typeof item[k] === 'object' && item[k] !== null ? JSON.stringify(item[k]) : item[k]);
          await q(`INSERT INTO \`${ent}\` (${keys.map(k => '`' + k + '`').join(',')}, created_at, updated_at) VALUES (${keys.map(() => '?').join(',')}, NOW(), NOW())`, vals);
          totalImported++;
        }
      }
    }

    res.json({ success: true, count: totalImported });
  } catch (err) {
    res.status(500).json({ message: 'JSON import error: ' + err.message });
  }
});

app.post('/api/upload', auth, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ message: 'No file received' });
  const ext = path.extname(req.file.originalname || '');
  const final = req.file.path + ext;
  fs.renameSync(req.file.path, final);
  const url = `/uploads/${path.basename(final)}`;
  res.json({ url, name: req.file.originalname, size: req.file.size });
});

// Storage & Archives Endpoints
app.get('/api/storage', auth, async (req, res) => {
  try {
    const { category } = req.query;
    let sql = 'SELECT * FROM storage_archives WHERE record_status != "archived"';
    const params = [];
    if (category && category !== 'all') {
      sql += ' AND category = ?';
      params.push(category);
    }
    sql += ' ORDER BY id DESC LIMIT 500';
    const rows = await q(sql, params);
    const parsed = rows.map(r => {
      let meta = {};
      if (typeof r.meta_json === 'string') {
        try { meta = JSON.parse(r.meta_json); } catch {}
      } else if (r.meta_json && typeof r.meta_json === 'object') {
        meta = r.meta_json;
      }
      return { ...r, meta };
    });
    res.json(parsed);
  } catch (err) {
    console.error('Storage list error:', err);
    res.status(500).json({ message: 'Storage list error: ' + err.message });
  }
});

app.post('/api/storage', auth, notViewer, async (req, res) => {
  try {
    const { category = 'other', title, filename, file_url = '', file_size = '—', format = 'other', meta_json = {} } = req.body;
    if (!title || !filename) {
      return res.status(400).json({ message: 'Title and filename are required' });
    }
    const metaStr = typeof meta_json === 'string' ? meta_json : JSON.stringify(meta_json || {});
    const r = await q(
      'INSERT INTO storage_archives (category, title, filename, file_url, file_size, format, meta_json, record_status, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, "active", NOW(), NOW())',
      [category, title, filename, file_url, file_size, format.toLowerCase(), metaStr]
    );
    const item = (await q('SELECT * FROM storage_archives WHERE id = ?', [r.insertId]))[0];
    let meta = {};
    try { meta = JSON.parse(item.meta_json); } catch {}
    res.json({ ...item, meta });
  } catch (err) {
    console.error('Storage create error:', err);
    res.status(500).json({ message: 'Storage create error: ' + err.message });
  }
});

app.post('/api/storage/upload', auth, notViewer, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ message: 'No file received' });
    const originalName = req.file.originalname || 'document';
    const ext = path.extname(originalName).toLowerCase();
    const final = req.file.path + ext;
    fs.renameSync(req.file.path, final);
    const fileUrl = `/uploads/${path.basename(final)}`;

    // Calculate readable size
    const bytes = req.file.size || 0;
    let readableSize = '1 KB';
    if (bytes >= 1024 * 1024) readableSize = (bytes / (1024 * 1024)).toFixed(1) + ' MB';
    else if (bytes >= 1024) readableSize = Math.round(bytes / 1024) + ' KB';
    else readableSize = bytes + ' B';

    // Guess format and category
    let format = ext.replace('.', '') || 'bin';
    let category = req.body.category || 'other';
    if (!req.body.category) {
      if (['pptx', 'ppt'].includes(format)) category = 'ppt';
      else if (['xlsx', 'xls', 'csv'].includes(format)) category = 'excel';
      else if (['json'].includes(format)) category = 'occupancy';
    }

    const title = req.body.title || originalName.replace(/\.[^/.]+$/, '').replace(/[_-]/g, ' ');
    let meta = {};
    try {
      if (req.body.meta_json) meta = typeof req.body.meta_json === 'string' ? JSON.parse(req.body.meta_json) : req.body.meta_json;
    } catch {}

    const r = await q(
      'INSERT INTO storage_archives (category, title, filename, file_url, file_size, format, meta_json, record_status, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, "active", NOW(), NOW())',
      [category, title, originalName, fileUrl, readableSize, format, JSON.stringify(meta)]
    );
    const item = (await q('SELECT * FROM storage_archives WHERE id = ?', [r.insertId]))[0];
    res.json({ ...item, meta });
  } catch (err) {
    console.error('Storage upload error:', err);
    res.status(500).json({ message: 'Storage upload error: ' + err.message });
  }
});

app.post('/api/storage/snapshot-occupancy', auth, notViewer, async (req, res) => {
  try {
    const sites = await q('SELECT * FROM sites WHERE record_status = "active"');
    const campaigns = await q('SELECT * FROM campaigns WHERE record_status = "active"');
    const totalSites = sites.length || 22;
    
    // Live campaigns running right now
    const now = new Date();
    const activeCampaigns = campaigns.filter(c => {
      const start = c.start_date ? new Date(c.start_date) : null;
      const end = c.end_date ? new Date(c.end_date) : null;
      return (!start || start <= now) && (!end || end >= now);
    });

    const occupiedCodes = new Set();
    let totalRevenue = 0;
    activeCampaigns.forEach(c => {
      if (c.site_code) occupiedCodes.add(c.site_code.toUpperCase());
      totalRevenue += Number(c.total_amount || c.revenue || 0);
    });

    const occupiedCount = Math.max(occupiedCodes.size, sites.filter(s => ['occupied', 'booked', 'active'].includes(String(s.availability || '').toLowerCase())).length);
    const vacantCount = Math.max(0, totalSites - occupiedCount);
    const occupancyPct = totalSites > 0 ? Math.round((occupiedCount / totalSites) * 100) : 0;

    const monthLabel = req.body.month || now.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    const title = `${monthLabel} Site Occupancy Snapshot`;
    const filename = `Occupancy_${monthLabel.replace(/\s+/g, '_')}.json`;

    const meta = {
      month: monthLabel,
      totalSites,
      occupiedSites: occupiedCount,
      vacantSites: vacantCount,
      occupancyPct,
      revenue: totalRevenue,
      activeCampaignsCount: activeCampaigns.length,
      timestamp: new Date().toISOString()
    };

    const r = await q(
      'INSERT INTO storage_archives (category, title, filename, file_url, file_size, format, meta_json, record_status, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, "active", NOW(), NOW())',
      ['occupancy', title, filename, '', '12 KB', 'json', JSON.stringify(meta)]
    );
    const item = (await q('SELECT * FROM storage_archives WHERE id = ?', [r.insertId]))[0];
    res.json({ ...item, meta });
  } catch (err) {
    console.error('Storage snapshot error:', err);
    res.status(500).json({ message: 'Snapshot error: ' + err.message });
  }
});

app.post('/api/storage/batch-delete', auth, managerOrAdmin, async (req, res) => {
  try {
    const { ids } = req.body || {};
    if (!Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ message: 'No IDs provided for batch deletion' });
    }
    const cleanIds = ids.map(Number).filter(n => !isNaN(n) && n > 0);
    if (!cleanIds.length) return res.status(400).json({ message: 'Invalid IDs' });

    const items = await q(`SELECT id, file_url FROM storage_archives WHERE id IN (${cleanIds.map(() => '?').join(',')})`, cleanIds);
    for (const item of items) {
      if (item.file_url && item.file_url.startsWith('/uploads/')) {
        const p = path.resolve(uploadDir, path.basename(item.file_url));
        if (fs.existsSync(p)) {
          try { fs.unlinkSync(p); } catch {}
        }
      }
    }

    await q(`DELETE FROM storage_archives WHERE id IN (${cleanIds.map(() => '?').join(',')})`, cleanIds);
    res.json({ success: true, count: cleanIds.length });
  } catch (err) {
    console.error('Storage batch-delete error:', err);
    res.status(500).json({ message: 'Batch delete error: ' + err.message });
  }
});

app.delete('/api/storage/:id', auth, managerOrAdmin, async (req, res) => {
  try {
    const item = (await q('SELECT * FROM storage_archives WHERE id = ?', [req.params.id]))[0];
    if (!item) return res.status(404).json({ message: 'Record not found' });
    await q('DELETE FROM storage_archives WHERE id = ?', [req.params.id]);
    if (item.file_url && item.file_url.startsWith('/uploads/')) {
      const p = path.resolve(uploadDir, path.basename(item.file_url));
      if (fs.existsSync(p)) {
        try { fs.unlinkSync(p); } catch {}
      }
    }
    res.json({ success: true, id: Number(req.params.id) });
  } catch (err) {
    console.error('Storage delete error:', err);
    res.status(500).json({ message: 'Delete error: ' + err.message });
  }
});

// Generic CRUD endpoints for entities
app.get('/api/:entity', auth, async (req, res) => {
  try {
    const e = safeEntity(req, res);
    if (!e) return;
    const limit = Math.min(Number(req.query.limit || 1000), 5000);
    let rows;
    if (e.table === 'electricity') {
      rows = await q(`
        SELECT 
          e.*,
          COALESCE(NULLIF(e.site_code,''), s.site_code, '') AS site_code,
          COALESCE(NULLIF(e.location,''), NULLIF(s.address,''), s.area, s.city, '') AS location,
          COALESCE(NULLIF(e.size,''), s.size, '') AS size,
          COALESCE(NULLIF(e.meter_no,''), s.meter_no, '') AS meter_no,
          COALESCE(NULLIF(e.service_number,''), '') AS service_number,
          COALESCE(NULLIF(e.t_number,''), '') AS t_number,
          COALESCE(NULLIF(e.bill_type,''), '') AS bill_type,
          COALESCE(NULLIF(e.amount, 0), e.payment_amount, 0) AS amount,
          COALESCE(NULLIF(e.payment_amount, 0), e.amount, 0) AS payment_amount
        FROM electricity e
        LEFT JOIN sites s ON (e.site_id = s.id OR (e.site_code != '' AND e.site_code = s.site_code))
        WHERE e.record_status = 'active'
        ORDER BY e.id DESC
        LIMIT ${limit}
      `);
    } else {
      const cols = await tableColumns(e.table);
      if (cols.has('record_status')) {
        rows = await q(`SELECT * FROM \`${e.table}\` WHERE record_status != 'archived' ORDER BY ${e.order} LIMIT ${limit}`);
      } else {
        rows = await q(`SELECT * FROM \`${e.table}\` ORDER BY ${e.order} LIMIT ${limit}`);
      }
    }
    
    // Auto-seed sites if table is empty
    if (e.table === 'sites' && (!rows || rows.length === 0)) {
      const seedCandidates = [
        path.resolve(__dirname, 'data/all-embedded-data.json'),
        path.resolve(__dirname, '../data/all-embedded-data.json'),
        path.resolve(process.cwd(), 'server/src/data/all-embedded-data.json'),
        path.resolve(process.cwd(), 'sql/all-embedded-data.json')
      ];
      const seedPath = seedCandidates.find(p => fs.existsSync(p));
      if (seedPath) {
        try {
          const raw = JSON.parse(fs.readFileSync(seedPath, 'utf8'));
          if (Array.isArray(raw.sites)) {
            for (const s of raw.sites) {
              const code = String(s.site_code || '').trim();
              if (!code) continue;
              const parts = (s.gps || '').split(',').map(x => parseFloat(x.trim()));
              const lat = isNaN(parts[0]) ? null : parts[0];
              const lng = isNaN(parts[1]) ? null : parts[1];
              const flags = typeof s.flags === 'string' ? s.flags : JSON.stringify(s.flags || {});
              await q(`INSERT INTO sites (
                site_code, city, area, address, size, media_type, lighting, facing,
                ownership, availability, vendor_name, meter_no, monthly_cost, monthly_rate,
                latitude, longitude, gps, notes, flags, record_status, created_at, updated_at
              ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?, 'active', NOW(), NOW())
              ON DUPLICATE KEY UPDATE site_code=VALUES(site_code)`, [
                code, s.city || 'Ahmedabad', s.area || '', s.address || '', s.size || '',
                s.media_type || 'Hoarding', s.lighting || 'BL', s.facing || '', s.ownership || 'Owned',
                s.availability || 'Available', s.vendor_name || '', s.meter_no || '',
                Number(s.monthly_cost || 0), Number(s.monthly_rate || 0), lat, lng, s.gps || '', s.notes || '', flags
              ]);
            }
            rows = await q(`SELECT * FROM \`sites\` ORDER BY id DESC LIMIT ${limit}`);
          }
        } catch (seedErr) {
          console.warn('Auto-seed in route notice:', seedErr.message);
        }
      }
    }

    for (const r of rows) {
      for (const k of ['flags', 'images_json']) {
        if (typeof r[k] === 'string') {
          try { r[k] = JSON.parse(r[k]); } catch {}
        }
      }
    }
    res.json(rows);
  } catch (err) {
    console.error(`Error loading ${req.params.entity}:`, err.message);
    res.status(500).json({ message: 'Database error: ' + err.message });
  }
});

app.get('/api/:entity/:id', auth, async (req, res) => {
  const e = safeEntity(req, res);
  if (!e) return;
  const rows = await q(`SELECT * FROM \`${e.table}\` WHERE id=? LIMIT 1`, [req.params.id]);
  if (!rows[0]) return res.status(404).json({ message: 'Record not found' });
  res.json(rows[0]);
});

app.post('/api/:entity', auth, notViewer, async (req, res) => {
  try {
    const e = safeEntity(req, res);
    if (!e) return;
    if (['proposals', 'invoices', 'clients'].includes(req.params.entity) && !['admin', 'manager'].includes(req.user?.role)) {
      return res.status(403).json({ message: 'Manager or Administrator access required' });
    }
    const data = await cleanData(e.table, req.body);
    if (req.params.entity === 'electricity') {
      if (data.amount && !data.payment_amount) data.payment_amount = data.amount;
      if (data.payment_amount && !data.amount) data.amount = data.payment_amount;
      if (!data.due_date) data.due_date = new Date().toISOString().slice(0, 10);
      if (data.site_code || data.site_id) {
        const siteRows = await q('SELECT * FROM sites WHERE id=? OR site_code=? LIMIT 1', [data.site_id || 0, data.site_code || '']);
        if (siteRows[0]) {
          const s = siteRows[0];
          if (!data.site_code) data.site_code = s.site_code;
          if (!data.site_id) data.site_id = s.id;
          if (!data.location) data.location = s.address || s.area || s.city || '';
          if (!data.size) data.size = s.size || '';
          if (!data.meter_no) data.meter_no = s.meter_no || '';
        }
      }
    }
    if (req.params.entity === 'campaigns') {
      if (!data.booking_code) data.booking_code = `MB-BK-${Date.now()}`;
      if (!data.start_date) data.start_date = new Date().toISOString().slice(0, 10);
      if (!data.end_date) {
        const d = new Date();
        d.setDate(d.getDate() + 30);
        data.end_date = d.toISOString().slice(0, 10);
      }
    }
    const keys = Object.keys(data);
    if (!keys.length) return res.status(400).json({ message: 'No valid fields' });
    const sql = `INSERT INTO \`${e.table}\` (${keys.map(k => '`' + k + '`').join(',')},created_at,updated_at) VALUES (${keys.map(() => '?').join(',')},NOW(),NOW())`;
    const result = await q(sql, keys.map(k => data[k]));
    await audit(req.user, 'Created', req.params.entity, result.insertId, null, data);
    // Auto-sync linked campaigns and site availability after any campaign change
    if (req.params.entity === 'campaigns') {
      await syncLinkedCampaigns(q, result.insertId);
      await syncSiteAvailability();
    }
    const rows = await q(`SELECT * FROM \`${e.table}\` WHERE id=?`, [result.insertId]);
    res.status(201).json(rows[0]);
  } catch (err) {
    console.error('Error creating entity ' + req.params.entity, err.message);
    res.status(500).json({ message: err.message });
  }
});

app.put('/api/:entity/:id', auth, notViewer, async (req, res) => {
  const e = safeEntity(req, res);
  if (!e) return;
  if (['proposals', 'invoices', 'clients'].includes(req.params.entity) && !['admin', 'manager'].includes(req.user?.role)) {
    return res.status(403).json({ message: 'Manager or Administrator access required' });
  }
  const before = (await q(`SELECT * FROM \`${e.table}\` WHERE id=?`, [req.params.id]))[0];
  if (!before) return res.status(404).json({ message: 'Record not found' });
  const data = await cleanData(e.table, req.body);
  if (req.params.entity === 'electricity') {
    if (data.amount && !data.payment_amount) data.payment_amount = data.amount;
    if (data.payment_amount && !data.amount) data.amount = data.payment_amount;
    if (data.site_code || data.site_id) {
      const siteRows = await q('SELECT * FROM sites WHERE id=? OR site_code=? LIMIT 1', [data.site_id || 0, data.site_code || '']);
      if (siteRows[0]) {
        const s = siteRows[0];
        if (!data.site_code && !before.site_code) data.site_code = s.site_code;
        if (!data.location && !before.location) data.location = s.address || s.area || s.city || '';
        if (!data.size && !before.size) data.size = s.size || '';
        if (!data.meter_no && !before.meter_no) data.meter_no = s.meter_no || '';
      }
    }
  }
  if (req.params.entity === 'campaigns') {
    if (data.total_amount && !data.revenue) data.revenue = data.total_amount;
    if (data.revenue && !data.total_amount) data.total_amount = data.revenue;
    if (data.display && !data.campaign_name) data.campaign_name = data.display;
    if (data.campaign_name && !data.display) data.display = data.campaign_name;
  }
  const keys = Object.keys(data);
  if (keys.length) await q(`UPDATE \`${e.table}\` SET ${keys.map(k => '`' + k + '`=?').join(',')},updated_at=NOW() WHERE id=?`, [...keys.map(k => data[k]), req.params.id]);
  const after = (await q(`SELECT * FROM \`${e.table}\` WHERE id=?`, [req.params.id]))[0];
  await audit(req.user, 'Updated', req.params.entity, req.params.id, before, after);
  // Auto-sync linked campaigns and site availability after any campaign update
  if (req.params.entity === 'campaigns') {
    await syncLinkedCampaigns(q, req.params.id);
    await syncSiteAvailability();
  }
  res.json(after);
});

app.delete('/api/:entity/:id', auth, managerOrAdmin, async (req, res) => {
  const e = safeEntity(req, res);
  if (!e) return;
  const cols = await tableColumns(e.table);
  const isHardDelete = e.table === 'campaigns' || e.table === 'electricity' || req.query.hard === 'true' || !cols.has('record_status');
  if (isHardDelete) {
    await q(`DELETE FROM \`${e.table}\` WHERE id=?`, [req.params.id]);
    await audit(req.user, 'Deleted', req.params.entity, req.params.id);
  } else {
    await q(`UPDATE \`${e.table}\` SET record_status='archived',updated_at=NOW() WHERE id=?`, [req.params.id]);
    await audit(req.user, 'Archived', req.params.entity, req.params.id);
  }
  // Auto-delete linked campaigns and sync site availability after any campaign deletion
  if (req.params.entity === 'campaigns') {
    await deleteLinkedCampaigns(q, req.params.id);
    await syncSiteAvailability();
  }
  res.json({ success: true });
});

app.post('/api/:entity/batch-delete', auth, managerOrAdmin, async (req, res) => {
  const e = safeEntity(req, res);
  if (!e) return;
  const { ids, hard } = req.body || {};
  if (!Array.isArray(ids) || ids.length === 0) {
    return res.status(400).json({ message: 'No IDs provided for batch deletion' });
  }
  const cleanIds = ids.map(x => Number(x)).filter(x => !isNaN(x) && x > 0);
  if (cleanIds.length === 0) {
    return res.status(400).json({ message: 'Invalid IDs for batch deletion' });
  }
  const cols = await tableColumns(e.table);
  const isHardDelete = e.table === 'campaigns' || e.table === 'electricity' || hard === true || req.query.hard === 'true' || !cols.has('record_status');
  
  const placeholders = cleanIds.map(() => '?').join(',');
  if (isHardDelete) {
    await q(`DELETE FROM \`${e.table}\` WHERE id IN (${placeholders})`, cleanIds);
    for (const id of cleanIds) {
      await audit(req.user, 'Deleted', req.params.entity, id);
    }
  } else {
    await q(`UPDATE \`${e.table}\` SET record_status='archived', updated_at=NOW() WHERE id IN (${placeholders})`, cleanIds);
    for (const id of cleanIds) {
      await audit(req.user, 'Archived', req.params.entity, id);
    }
  }
  if (req.params.entity === 'campaigns') {
    await batchDeleteLinkedCampaigns(q, cleanIds);
    await syncSiteAvailability();
  }
  res.json({ success: true, count: cleanIds.length });
});

// Linked site conflict check & manual sync endpoints
app.get('/api/sites/conflicts/:code', auth, (req, res) => {
  res.json(getConflictSummary(req.params.code) || {});
});
app.post('/api/sites/sync-availability', auth, managerOrAdmin, async (req, res) => {
  await syncSiteAvailability();
  res.json({ success: true, message: 'All site availabilities synchronized.' });
});
const distCandidates = [
  path.resolve(__dirname, '../../client/dist'),
  path.resolve(__dirname, '../client/dist'),
  path.resolve(process.cwd(), 'client/dist'),
  path.resolve(process.cwd(), 'dist')
];
const clientDist = distCandidates.find(p => fs.existsSync(p));
if (clientDist) {
  console.log(`Serving static frontend from: ${clientDist}`);
  app.use(express.static(clientDist));
  app.get('*', (req, res) => res.sendFile(path.join(clientDist, 'index.html')));
} else {
  console.warn('Frontend build dist folder not found. API routes are active.');
  app.get('/', (req, res) => res.json({ status: 'API is running', endpoints: '/api/health' }));
}

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Site Control API & Web App listening on port ${PORT}`);
});

async function initDb() {
  try {
    await pool.query('SELECT 1');
    console.log('Database connected successfully');

    // Automatically create tables from schema.sql if they do not exist
    const schemaCandidates = [
      path.resolve(__dirname, '../../sql/schema.sql'),
      path.resolve(__dirname, '../sql/schema.sql'),
      path.resolve(process.cwd(), 'sql/schema.sql'),
      path.resolve(process.cwd(), '../sql/schema.sql')
    ];
    const schemaPath = schemaCandidates.find(p => fs.existsSync(p));
    if (schemaPath) {
      try {
        const sql = fs.readFileSync(schemaPath, 'utf8');
        const stmts = sql.split(';').map(s => s.trim()).filter(s => s.length > 0);
        for (const stmt of stmts) {
          try {
            await pool.query(stmt);
          } catch (e) {
            // Ignore benign statements
          }
        }
        console.log('Database tables verified and ready.');
      } catch (schemaErr) {
        console.warn('Auto-schema execution notice:', schemaErr.message);
      }
    }

    try {
      const emails = await q('SELECT COUNT(*) c FROM users');
      if (!emails[0].c && process.env.ADMIN_EMAIL && process.env.ADMIN_PASSWORD) {
        const hash = await bcrypt.hash(process.env.ADMIN_PASSWORD, 12);
        await q('INSERT INTO users(name,email,password_hash,role,status,created_at,updated_at) VALUES(?,?,?,?,"active",NOW(),NOW())', ['Administrator', process.env.ADMIN_EMAIL, hash, 'admin']);
        console.log('Initial admin user created');
      }
    } catch (tblErr) {
      console.warn('User table check notice:', tblErr.message);
    }

    try {
      await q('ALTER TABLE electricity MODIFY due_date DATE NULL DEFAULT NULL');
      await q('ALTER TABLE campaigns MODIFY start_date DATE NULL DEFAULT NULL, MODIFY end_date DATE NULL DEFAULT NULL');
    } catch (_) {}

    await ensureCampaignColumns();

    // Auto-seed initial Media Buzz sites & settings if empty
    try {
      const siteCount = await q('SELECT COUNT(*) c FROM sites WHERE record_status="active"');
      if (!siteCount[0]?.c) {
        const seedCandidates = [
          path.resolve(__dirname, 'data/all-embedded-data.json'),
          path.resolve(__dirname, '../data/all-embedded-data.json'),
          path.resolve(process.cwd(), 'server/src/data/all-embedded-data.json'),
          path.resolve(process.cwd(), 'sql/all-embedded-data.json')
        ];
        const seedPath = seedCandidates.find(p => fs.existsSync(p));
        if (seedPath) {
          const raw = JSON.parse(fs.readFileSync(seedPath, 'utf8'));
          if (Array.isArray(raw.sites)) {
            for (const s of raw.sites) {
              const code = String(s.site_code || '').trim();
              if (!code) continue;
              const parts = (s.gps || '').split(',').map(x => parseFloat(x.trim()));
              const lat = isNaN(parts[0]) ? null : parts[0];
              const lng = isNaN(parts[1]) ? null : parts[1];
              const flags = typeof s.flags === 'string' ? s.flags : JSON.stringify(s.flags || {});
              await q(`INSERT INTO sites (
                site_code, city, area, address, size, media_type, lighting, facing,
                ownership, availability, vendor_name, meter_no, monthly_cost, monthly_rate,
                latitude, longitude, gps, notes, flags, record_status, created_at, updated_at
              ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?, 'active', NOW(), NOW())
              ON DUPLICATE KEY UPDATE site_code=VALUES(site_code)`, [
                code,
                s.city || 'Ahmedabad',
                s.area || '',
                s.address || '',
                s.size || '',
                s.media_type || 'Hoarding',
                s.lighting || 'BL',
                s.facing || '',
                s.ownership || 'Owned',
                s.availability || 'Available',
                s.vendor_name || '',
                s.meter_no || '',
                Number(s.monthly_cost || 0),
                Number(s.monthly_rate || 0),
                lat,
                lng,
                s.gps || '',
                s.notes || '',
                flags
              ]);
            }
            console.log(`Seeded ${raw.sites.length} initial Media Buzz sites successfully.`);
          }

          if (raw.default_settings) {
            for (const [k, v] of Object.entries(raw.default_settings)) {
              await q('INSERT INTO settings (setting_key, setting_value, updated_at) VALUES (?,?,NOW()) ON DUPLICATE KEY UPDATE setting_value=VALUES(setting_value)', [k, String(v ?? '')]);
            }
            console.log('Seeded default Media Buzz settings successfully.');
          }
        }
      }

      // Seed Clients if empty
      const clientCount = await q('SELECT COUNT(*) c FROM clients');
      if (!clientCount[0]?.c) {
        await q(`INSERT INTO clients (client_name, company, primary_contact, email, phone, billing_address, gst_number, status, created_at, updated_at) VALUES
          ('Rajyash Group', 'Rajyash Estates Pvt Ltd', 'Pratik Patel', 'contact@rajyash.com', '+91 98250 11223', 'Rajyash House, Ambli-Bopal Road, Ahmedabad', '24AAACR1234F1Z5', 'active', NOW(), NOW()),
          ('Adani Realty', 'Adani Infrastructure Developers', 'Vikram Mehta', 'ooh@adani.com', '+91 79 2656 5555', 'Adani Corporate House, Shantigram, Ahmedabad', '24AAACA5678B1Z2', 'active', NOW(), NOW()),
          ('Zydus Healthcare', 'Zydus Lifesciences Ltd', 'Sneha Dave', 'media@zyduslife.com', '+91 79 4804 0000', 'Zydus Corporate Park, SG Highway, Ahmedabad', '24AAACZ9988G1Z9', 'active', NOW(), NOW()),
          ('Havmor Ice Cream', 'Havmor Foods Ltd', 'Amit Shah', 'marketing@havmor.com', '+91 79 2642 1100', 'Commerce House, Navrangpura, Ahmedabad', '24AAACH4433H1Z1', 'active', NOW(), NOW()),
          ('Iscon Group', 'JP Iscon Builders', 'Rajesh Agarwal', 'info@iscongroup.com', '+91 98795 44332', 'Iscon Elegance, Prahladnagar, Ahmedabad', '24AAACI7766K1Z4', 'active', NOW(), NOW())`);
        console.log('Seeded initial clients.');
      }

      // Seed Vendors if empty
      const vendorCount = await q('SELECT COUNT(*) c FROM vendors');
      if (!vendorCount[0]?.c) {
        await q(`INSERT INTO vendors (name, service, contact_person, phone, email, cities, rating, status, created_at, updated_at) VALUES
          ('Gujarat Printers & Signage', 'Printing', 'Mukesh Bhai', '+91 98240 55441', 'mukesh@gujaratprinters.com', 'Ahmedabad, Gandhinagar, Surat', 4.85, 'active', NOW(), NOW()),
          ('Ahmedabad Neon & LED Arts', 'Mounting & Electrical', 'Haresh Solanki', '+91 98980 12345', 'info@ahmedabadneon.com', 'Ahmedabad, Vadodara', 4.90, 'active', NOW(), NOW()),
          ('Om Sai Mounting Services', 'Mounting', 'Ramesh Parmar', '+91 94260 88776', 'omsaimounting@gmail.com', 'Ahmedabad, Rajkot', 4.65, 'active', NOW(), NOW()),
          ('Apex Outdoor Fabricators', 'Fabrication & Maintenance', 'Ketan Shah', '+91 98255 33221', 'ketan@apexfab.in', 'Ahmedabad', 4.75, 'active', NOW(), NOW())`);
        console.log('Seeded initial vendors.');
      }

      // Seed Campaigns if empty
      const campCount = await q('SELECT COUNT(*) c FROM campaigns');
      if (!campCount[0]?.c) {
        const sRows = await q('SELECT id, site_code FROM sites LIMIT 5');
        const s1 = sRows[0]?.id || 1, code1 = sRows[0]?.site_code || 'AMD-GT-001';
        const s2 = sRows[1]?.id || 2, code2 = sRows[1]?.site_code || 'AMD-UP-002';
        const s3 = sRows[2]?.id || 3, code3 = sRows[2]?.site_code || 'AMD-HD-005';

        await q(`INSERT INTO campaigns (
          booking_code, site_id, site_code, client, brand, campaign_name,
          booking_date, start_date, end_date, mounting_date, printing_status, mounting_status,
          validation_15_date, final_validation_date, revenue, vendor_cost, printing_cost, mounting_cost,
          electricity_cost, other_cost, invoice_status, hard_copy_status, notes, record_status, created_at, updated_at
        ) VALUES
          ('MB-BK-2026-001', ?, ?, 'Rajyash Group', 'Rajyash Estates', 'Diwali Launch Ahmedabad', CURDATE(), CURDATE(), DATE_ADD(CURDATE(), INTERVAL 30 DAY), NULL, 'Completed', 'Pending', DATE_ADD(CURDATE(), INTERVAL 15 DAY), DATE_ADD(CURDATE(), INTERVAL 30 DAY), 350000, 45000, 32000, 18000, 8500, 2000, 'Pending', 'Pending', 'Prime gantry Diwali campaign', 'active', NOW(), NOW()),
          ('MB-BK-2026-002', ?, ?, 'Adani Realty', 'Adani Shantigram', 'Township Phase 2 Launch', DATE_SUB(CURDATE(), INTERVAL 24 DAY), DATE_SUB(CURDATE(), INTERVAL 20 DAY), DATE_ADD(CURDATE(), INTERVAL 5 DAY), DATE_SUB(CURDATE(), INTERVAL 19 DAY), 'Completed', 'Mounted', DATE_SUB(CURDATE(), INTERVAL 5 DAY), DATE_ADD(CURDATE(), INTERVAL 5 DAY), 420000, 52000, 38000, 22000, 9200, 3000, 'Sent', 'Dispatched', '15-day validation complete', 'active', NOW(), NOW()),
          ('MB-BK-2026-003', ?, ?, 'Zydus Healthcare', 'Zydus Wellness', 'Health First Hoardings', DATE_SUB(CURDATE(), INTERVAL 16 DAY), DATE_SUB(CURDATE(), INTERVAL 15 DAY), DATE_ADD(CURDATE(), INTERVAL 15 DAY), DATE_SUB(CURDATE(), INTERVAL 14 DAY), 'Completed', 'Mounted', CURDATE(), DATE_ADD(CURDATE(), INTERVAL 15 DAY), 280000, 35000, 26000, 15000, 6800, 1500, 'Pending', 'Pending', '15-day validation photo due today', 'active', NOW(), NOW())`,
          [s1, code1, s2, code2, s3, code3]
        );
        console.log('Seeded initial campaigns.');
      }

      // Seed Electricity if empty
      const elecCount = await q('SELECT COUNT(*) c FROM electricity');
      if (!elecCount[0]?.c) {
        await q(`INSERT INTO electricity (
          site_code, location, meter_no, size, service_number, t_number, bill_type,
          payment_amount, billing_month, bill_date, due_date, units, rate, amount,
          payment_status, paid_date, payment_reference, notes, record_status, created_at, updated_at
        ) VALUES
          ('AMD-GT-001', 'Shivranjani Cross Roads, Ahmedabad', 'MTR-UGVCL-8841', '30x10 ft', 'SRV-998241', 'T-4401', 'UGVCL', 7850, 'Aug 2026', DATE_SUB(CURDATE(), INTERVAL 18 DAY), DATE_SUB(CURDATE(), INTERVAL 2 DAY), 850, 9.23, 7850, 'Pending', NULL, '', 'Overdue meter bill', 'active', NOW(), NOW()),
          ('AMD-UP-002', 'SG Highway near YMCA Club, Ahmedabad', 'MTR-TORRENT-3312', '40x20 ft', 'SRV-887412', 'T-5512', 'Torrent Power', 12400, 'Aug 2026', DATE_SUB(CURDATE(), INTERVAL 10 DAY), DATE_ADD(CURDATE(), INTERVAL 6 DAY), 1320, 9.39, 12400, 'Pending', NULL, '', 'Due soon', 'active', NOW(), NOW()),
          ('AMD-HD-005', 'Sindhubhavan Road, Ahmedabad', 'MTR-TORRENT-1198', '50x20 ft', 'SRV-665209', 'T-2209', 'Torrent Power', 15800, 'Jul 2026', DATE_SUB(CURDATE(), INTERVAL 40 DAY), DATE_SUB(CURDATE(), INTERVAL 25 DAY), 1680, 9.40, 15800, 'Paid', DATE_SUB(CURDATE(), INTERVAL 28 DAY), 'UPI-9923847291', 'Paid on time', 'active', NOW(), NOW())`);
        console.log('Seeded initial electricity bills.');
      }

      // Seed Invoices if empty
      const invCount = await q('SELECT COUNT(*) c FROM invoices');
      if (!invCount[0]?.c) {
        await q(`INSERT INTO invoices (
          campaign_id, client_id, requested_date, invoice_no, invoice_date, invoice_amount,
          invoice_status, hard_copy_required, hard_copy_status, courier_name, tracking_number,
          dispatch_date, delivered_date, payment_status, notes, record_status, created_at, updated_at
        ) VALUES
          (2, 2, DATE_SUB(CURDATE(), INTERVAL 8 DAY), 'MB-INV-2026-088', DATE_SUB(CURDATE(), INTERVAL 7 DAY), 495600, 'Sent', 1, 'Dispatched', 'BlueDart', 'BD998234109IN', DATE_SUB(CURDATE(), INTERVAL 5 DAY), NULL, 'Pending', 'Adani Shantigram GST invoice', 'active', NOW(), NOW()),
          (1, 1, CURDATE(), 'MB-INV-2026-089', CURDATE(), 413000, 'Pending', 1, 'Pending', '', '', NULL, NULL, 'Pending', 'Rajyash Diwali advance invoice', 'active', NOW(), NOW())`);
        console.log('Seeded initial invoices.');
      }

      // Seed Notifications if empty
      const notifCount = await q('SELECT COUNT(*) c FROM notifications');
      if (!notifCount[0]?.c) {
        await q(`INSERT INTO notifications (user_id, type, object_type, title, message, is_read, fingerprint, created_at) VALUES
          (0, 'mounting_due', 'campaign', 'Mounting Overdue: AMD-GT-001', 'Rajyash Group campaign scheduled start date was reached but mounting is pending.', 0, 'notif-mount-001', NOW()),
          (0, 'validation_due', 'campaign', '15-Day Validation Due: AMD-HD-005', 'Zydus Healthcare campaign requires day/night inspection photo proof.', 0, 'notif-val-002', NOW()),
          (0, 'electricity_due', 'electricity', 'Electricity Bill Overdue: MTR-UGVCL-8841', 'Bill amount ₹7,850 for AMD-GT-001 is past due date.', 0, 'notif-elec-003', NOW())`);
        console.log('Seeded initial notifications.');
      }

      // Seed Storage Archives if empty
      const storageCount = await q('SELECT COUNT(*) c FROM storage_archives');
      if (!storageCount[0]?.c) {
        await q(`INSERT INTO storage_archives (category, title, filename, file_url, file_size, format, meta_json, record_status, created_at, updated_at) VALUES
          ('ppt', 'Diwali 2026 Prime Sites Pitch Deck', 'MediaBuzz_Automated-PPT_Diwali2026.pptx', '', '4.8 MB', 'pptx', '{"slides":8,"client":"All Prime Sites","sites":["AMD-GT-001","AMD-UP-002","AMD-HD-005"],"orientation":"16:9 Widescreen"}', 'active', DATE_SUB(NOW(), INTERVAL 2 DAY), NOW()),
          ('ppt', 'Rajyash Group - Ahmedabad Outdoor Showcase', 'Rajyash_Group_Outdoor_Showcase.pptx', '', '3.2 MB', 'pptx', '{"slides":5,"client":"Rajyash Group","sites":["AMD-GT-001","AMD-HD-005"]}', 'active', DATE_SUB(NOW(), INTERVAL 5 DAY), NOW()),
          ('ppt', 'Tata Motors EV Launch Presentation', 'Tata_Motors_EV_Launch_Deck.pptx', '', '2.9 MB', 'pptx', '{"slides":6,"client":"Tata Motors EV","sites":["AMD-UP-002"]}', 'active', DATE_SUB(NOW(), INTERVAL 9 DAY), NOW()),
          ('excel', 'Campaign Tracker Master Sheet (20 Columns)', 'MediaBuzz_Campaign_Tracker_Sep2026.xlsx', '', '380 KB', 'xlsx', '{"rows":24,"columns":20,"type":"Campaigns","month":"Sep 2026"}', 'active', DATE_SUB(NOW(), INTERVAL 1 DAY), NOW()),
          ('excel', 'Consolidated Electricity Bills (UGVCL & Torrent)', 'MediaBuzz_Electricity_Bills_Aug2026.xlsx', '', '210 KB', 'xlsx', '{"rows":18,"totalAmount":184500,"type":"Electricity"}', 'active', DATE_SUB(NOW(), INTERVAL 4 DAY), NOW()),
          ('excel', 'Master Sites Portfolio & Geo-Coordinates', 'MediaBuzz_Sites_Catalog_2026.xlsx', '', '512 KB', 'xlsx', '{"rows":22,"type":"Sites Catalog","cities":["Ahmedabad"]}', 'active', DATE_SUB(NOW(), INTERVAL 12 DAY), NOW()),
          ('occupancy', 'September 2026 Site Occupancy Snapshot', 'Occupancy_September_2026.json', '', '14 KB', 'json', '{"month":"September 2026","totalSites":22,"occupiedSites":19,"vacantSites":3,"occupancyPct":86,"revenue":1480000,"activeCampaignsCount":6}', 'active', DATE_SUB(NOW(), INTERVAL 8 DAY), NOW()),
          ('occupancy', 'August 2026 Site Occupancy Snapshot', 'Occupancy_August_2026.json', '', '14 KB', 'json', '{"month":"August 2026","totalSites":22,"occupiedSites":17,"vacantSites":5,"occupancyPct":77,"revenue":1290000,"activeCampaignsCount":5}', 'active', DATE_SUB(NOW(), INTERVAL 38 DAY), NOW()),
          ('occupancy', 'July 2026 Site Occupancy Snapshot', 'Occupancy_July_2026.json', '', '14 KB', 'json', '{"month":"July 2026","totalSites":22,"occupiedSites":16,"vacantSites":6,"occupancyPct":73,"revenue":1150000,"activeCampaignsCount":5}', 'active', DATE_SUB(NOW(), INTERVAL 69 DAY), NOW()),
          ('occupancy', 'June 2026 Site Occupancy Snapshot', 'Occupancy_June_2026.json', '', '14 KB', 'json', '{"month":"June 2026","totalSites":22,"occupiedSites":15,"vacantSites":7,"occupancyPct":68,"revenue":1020000,"activeCampaignsCount":4}', 'active', DATE_SUB(NOW(), INTERVAL 99 DAY), NOW())`);
        console.log('Seeded initial storage archives with generated PPTs, Excels, and historical occupancy snapshots.');
      }
      // Synchronize linked campaigns and site availability on startup
      await syncAllLinkedCampaigns(q);
      await syncSiteAvailability();
    } catch (seedErr) {
      console.warn('Site seed notice:', seedErr.message);
    }
  } catch (err) {
    console.error('Database connection warning (check DB_HOST, DB_USER, DB_PASSWORD, DB_NAME):', err.message);
  }
}

initDb();


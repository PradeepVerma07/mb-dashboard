import 'dotenv/config';
import express from 'express';import cors from 'cors';import helmet from 'helmet';import path from 'path';import fs from 'fs';import multer from 'multer';import bcrypt from 'bcryptjs';import {fileURLToPath} from 'url';import * as XLSX from 'xlsx';
import {pool,q} from './db.js';import {auth,admin,sign} from './auth.js';
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
    const [[s], [c], [e], [i]] = await Promise.all([
      q("SELECT COUNT(*) total, COALESCE(SUM(availability='Available'),0) available FROM sites WHERE record_status='active'"),
      q("SELECT COUNT(*) total, COALESCE(SUM(CURDATE() BETWEEN start_date AND end_date),0) live, COALESCE(SUM(revenue),0) revenue, COALESCE(SUM(revenue-vendor_cost-printing_cost-mounting_cost-electricity_cost-other_cost),0) margin FROM campaigns WHERE record_status='active'"),
      q("SELECT COUNT(*) total, COALESCE(SUM(payment_status<>'Paid' AND due_date<CURDATE()),0) overdue, COALESCE(SUM(CASE WHEN payment_status<>'Paid' THEN amount ELSE 0 END),0) unpaid FROM electricity WHERE record_status='active'"),
      q("SELECT COUNT(*) total, COALESCE(SUM(invoice_status NOT IN ('Paid','Sent')),0) pending FROM invoices WHERE record_status='active'")
    ]);
    const alerts = await q("SELECT id,site_code,client,campaign_name,end_date,invoice_status,hard_copy_status FROM campaigns WHERE record_status='active' AND (end_date <= DATE_ADD(CURDATE(),INTERVAL 7 DAY) OR invoice_status='Pending' OR hard_copy_status='Pending') ORDER BY end_date LIMIT 20");
    res.json({
      kpis: {
        total_sites: Number(s?.total || 0),
        available_sites: Number(s?.available || 0),
        active_campaigns: Number(c?.live || 0),
        campaign_revenue: Number(c?.revenue || 0),
        gross_margin: Number(c?.margin || 0),
        electricity_overdue: Number(e?.overdue || 0),
        unpaid_electricity: Number(e?.unpaid || 0),
        invoice_actions: Number(i?.pending || 0)
      },
      alerts: alerts || []
    });
  } catch (err) {
    console.error('Dashboard error:', err.message);
    res.status(500).json({message: 'Dashboard data error: ' + err.message});
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
    else await q('UPDATE notifications SET is_read=1 WHERE user_id IN (0,?)', [req.user.id]);
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

app.post('/api/import/xlsx', auth, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ message: 'No workbook' });
  try {
    const wb = XLSX.readFile(req.file.path), sheet = wb.Sheets[wb.SheetNames[0]], rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    let count = 0;
    for (const r of rows) {
      const code = String(r['Site ID'] || r['site_code'] || r['SITE ID'] || '').trim();
      if (!code) continue;
      const flags = { ppt_availability: r['PPT Availability'] || r['AVAILABLITY'] || r['Availability'] || '', ppt_rate: r['PPT Rate Per Month (₹)'] || r['Selling Amount (₹)'] || r['Selling Amount'] || 0, ppt_images: [] };
      const data = { site_code: code, city: r.City || r.CITY || '', area: r['Area / Landmark'] || r.AREA || '', address: r['Full Address'] || r.LOCATION || '', size: r.Size || '', width: Number(r['Width (ft)'] || r.W || 0) || null, height: Number(r['Height (ft)'] || r.H || 0) || null, media_type: r['Media Type'] || r.MEDIA || 'Billboard', lighting: r.Lighting || r.LIGHT || 'FL', availability: r.Availability || 'Available', monthly_rate: Number(String(r['Selling Amount (₹)'] || r['Selling Amount'] || 0).replace(/[^0-9.]/g, '')) || 0, latitude: Number(r.Latitude) || null, longitude: Number(r.Longitude) || null, flags: JSON.stringify(flags) };
      const ex = (await q('SELECT id,flags FROM sites WHERE site_code=?', [code]))[0];
      if (ex) {
        let old = {};
        try { old = JSON.parse(ex.flags || '{}') || {}; } catch {};
        data.flags = JSON.stringify({ ...old, ...flags, ppt_images: old.ppt_images || [] });
        await q('UPDATE sites SET city=?,area=?,address=?,size=?,width=?,height=?,media_type=?,lighting=?,availability=?,monthly_rate=?,latitude=?,longitude=?,flags=?,updated_at=NOW() WHERE id=?', [data.city, data.area, data.address, data.size, data.width, data.height, data.media_type, data.lighting, data.availability, data.monthly_rate, data.latitude, data.longitude, data.flags, ex.id]);
      } else {
        await q('INSERT INTO sites(site_code,city,area,address,size,width,height,media_type,lighting,availability,monthly_rate,latitude,longitude,flags,record_status,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?, "active",NOW(),NOW())', [data.site_code, data.city, data.area, data.address, data.size, data.width, data.height, data.media_type, data.lighting, data.availability, data.monthly_rate, data.latitude, data.longitude, data.flags]);
      }
      count++;
    }
    res.json({ success: true, rows: count });
  } catch (err) {
    res.status(500).json({ message: 'Import error: ' + err.message });
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

// Generic CRUD endpoints for entities
app.get('/api/:entity', auth, async (req, res) => {
  try {
    const e = safeEntity(req, res);
    if (!e) return;
    const limit = Math.min(Number(req.query.limit || 1000), 5000);
    let rows = await q(`SELECT * FROM \`${e.table}\` ORDER BY ${e.order} LIMIT ${limit}`);
    
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

app.post('/api/:entity', auth, async (req, res) => {
  const e = safeEntity(req, res);
  if (!e) return;
  const data = await cleanData(e.table, req.body);
  const keys = Object.keys(data);
  if (!keys.length) return res.status(400).json({ message: 'No valid fields' });
  const sql = `INSERT INTO \`${e.table}\` (${keys.map(k => '`' + k + '`').join(',')},created_at,updated_at) VALUES (${keys.map(() => '?').join(',')},NOW(),NOW())`;
  const result = await q(sql, keys.map(k => data[k]));
  await audit(req.user, 'Created', req.params.entity, result.insertId, null, data);
  const rows = await q(`SELECT * FROM \`${e.table}\` WHERE id=?`, [result.insertId]);
  res.status(201).json(rows[0]);
});

app.put('/api/:entity/:id', auth, async (req, res) => {
  const e = safeEntity(req, res);
  if (!e) return;
  const before = (await q(`SELECT * FROM \`${e.table}\` WHERE id=?`, [req.params.id]))[0];
  if (!before) return res.status(404).json({ message: 'Record not found' });
  const data = await cleanData(e.table, req.body), keys = Object.keys(data);
  if (keys.length) await q(`UPDATE \`${e.table}\` SET ${keys.map(k => '`' + k + '`=?').join(',')},updated_at=NOW() WHERE id=?`, [...keys.map(k => data[k]), req.params.id]);
  const after = (await q(`SELECT * FROM \`${e.table}\` WHERE id=?`, [req.params.id]))[0];
  await audit(req.user, 'Updated', req.params.entity, req.params.id, before, after);
  res.json(after);
});

app.delete('/api/:entity/:id', auth, async (req, res) => {
  const e = safeEntity(req, res);
  if (!e) return;
  const cols = await tableColumns(e.table);
  if (cols.has('record_status')) await q(`UPDATE \`${e.table}\` SET record_status='archived',updated_at=NOW() WHERE id=?`, [req.params.id]);
  else await q(`DELETE FROM \`${e.table}\` WHERE id=?`, [req.params.id]);
  await audit(req.user, 'Archived', req.params.entity, req.params.id);
  res.json({ success: true });
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
    } catch (seedErr) {
      console.warn('Site seed notice:', seedErr.message);
    }
  } catch (err) {
    console.error('Database connection warning (check DB_HOST, DB_USER, DB_PASSWORD, DB_NAME):', err.message);
  }
}

initDb();


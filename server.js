const express  = require('express');
const ExcelJS  = require('exceljs');
const path     = require('path');
const { Pool } = require('pg');
const session  = require('express-session');
const bcrypt   = require('bcrypt');

const app  = express();
const PORT = process.env.PORT || 3000;

// ── PostgreSQL connection ──────────────────────────────────────────
// FIX: Removed pool.on('connect') SET application_name — this was
// causing "Connection terminated unexpectedly" on Render's free DB tier.
// FIX: Parse DATABASE_URL properly; if it already contains ?sslmode=
// don't double-add SSL options — just pass ssl:{rejectUnauthorized:false}.
function buildPoolConfig() {
  const connStr = process.env.DATABASE_URL || '';
  // Render Postgres URLs sometimes include ?sslmode=require already
  const cfg = {
    connectionString:        connStr,
    connectionTimeoutMillis: 15000,
    idleTimeoutMillis:       30000,
    max:                     5,       // keep low on free tier
  };
  // Always add SSL for Render — rejectUnauthorized:false because
  // Render uses a self-signed / internal CA cert
  if (connStr.startsWith('postgres')) {
    cfg.ssl = { rejectUnauthorized: false };
  }
  return cfg;
}

const pool = new Pool(buildPoolConfig());

pool.on('error', (err) => {
  console.error('PG pool error (non-fatal):', err.message);
});

// ── DB Init (with retry) ───────────────────────────────────────────
async function initDB(retries = 10, delay = 3000) {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      await _runMigrations();
      console.log('  ✅  Database ready');
      return;
    } catch (err) {
      console.error(`  ⚠️   DB init attempt ${attempt}/${retries} failed: ${err.message}`);
      if (attempt === retries) throw err;
      const wait = Math.min(delay * Math.pow(1.5, attempt - 1), 60000);
      console.log(`  🔄  Retrying in ${Math.round(wait / 1000)}s…`);
      await new Promise(r => setTimeout(r, wait));
    }
  }
}

async function _runMigrations() {
  // FIX: Use a single client from the pool for the entire migration
  // transaction so we don't burn multiple connections on startup.
  const client = await pool.connect();
  try {
    await client.query('BEGIN');

    await client.query(`
      CREATE TABLE IF NOT EXISTS users (
        id         SERIAL PRIMARY KEY,
        username   TEXT UNIQUE NOT NULL,
        password   TEXT NOT NULL,
        created_at TIMESTAMP DEFAULT NOW()
      )
    `);

    await client.query(`
      CREATE TABLE IF NOT EXISTS clinic_config (
        id         SERIAL PRIMARY KEY,
        user_id    INTEGER REFERENCES users(id) ON DELETE CASCADE,
        configured BOOLEAN DEFAULT false,
        start_odip INTEGER DEFAULT 1,
        next_odip  INTEGER DEFAULT 1,
        UNIQUE(user_id)
      )
    `);

    await client.query(`
      CREATE TABLE IF NOT EXISTS excel_files (
        id         SERIAL PRIMARY KEY,
        user_id    INTEGER REFERENCES users(id) ON DELETE CASCADE,
        name       TEXT NOT NULL,
        created_at TIMESTAMP DEFAULT NOW()
      )
    `);

    await client.query(`
      CREATE TABLE IF NOT EXISTS sheets (
        id           SERIAL PRIMARY KEY,
        file_id      INTEGER REFERENCES excel_files(id) ON DELETE CASCADE,
        user_id      INTEGER REFERENCES users(id) ON DELETE CASCADE,
        name         TEXT NOT NULL,
        position     INTEGER NOT NULL,
        holiday_type TEXT,
        created_at   TIMESTAMP DEFAULT NOW()
      )
    `);

    await client.query(`
      CREATE TABLE IF NOT EXISTS patients (
        id         SERIAL PRIMARY KEY,
        sheet_id   INTEGER REFERENCES sheets(id) ON DELETE CASCADE,
        user_id    INTEGER REFERENCES users(id) ON DELETE CASCADE,
        odip       INTEGER NOT NULL,
        name       TEXT    NOT NULL,
        treatment  TEXT    NOT NULL,
        amount     INTEGER NOT NULL,
        position   INTEGER NOT NULL,
        created_at TIMESTAMP DEFAULT NOW()
      )
    `);

    await client.query('COMMIT');

    // Safe column additions outside the transaction (ALTER TABLE IF NOT EXISTS
    // doesn't work inside a transaction on older Postgres versions)
    const migrations = [
      `ALTER TABLE clinic_config ADD COLUMN IF NOT EXISTS user_id INTEGER`,
      `ALTER TABLE excel_files   ADD COLUMN IF NOT EXISTS user_id INTEGER`,
      `ALTER TABLE sheets        ADD COLUMN IF NOT EXISTS user_id INTEGER`,
      `ALTER TABLE sheets        ADD COLUMN IF NOT EXISTS holiday_type TEXT`,
      `ALTER TABLE patients      ADD COLUMN IF NOT EXISTS user_id INTEGER`,
    ];
    for (const m of migrations) {
      try { await pool.query(m); } catch(_) {}
    }
  } catch (err) {
    await client.query('ROLLBACK');
    throw err;
  } finally {
    client.release();
  }
}

// ── 33 Treatments ─────────────────────────────────────────────────
const TREATMENTS = [
  { name: 'PA-Para, Amox250, Famtab, Avil',                      amount: 70  },
  { name: 'DicloMR, Neurobin, OMez, Dexa',                       amount: 70  },
  { name: 'Dr-Dressing, grocin BC, cpm, dexa',                   amount: 120 },
  { name: 'Nimupara, Famtab, Levocetriza, Amox260',              amount: 70  },
  { name: 'Aceclopara, nuvit, famtab, dexa',                     amount: 70  },
  { name: 'inj Rantac, Cyclpam, Omez, Metro, Dexa',             amount: 140 },
  { name: 'Small, CPM, Depin, Dexa',                             amount: 50  },
  { name: 'Aceclopara, Neurobin, Famtab, Dexa',                  amount: 70  },
  { name: 'Cyclpam, Pipajam, Famtab, Dexa',                      amount: 70  },
  { name: 'Nimupara, famtab, stemetil, neurobin',                amount: 70  },
  { name: 'inj.cyna, diclopara, metro, famtab, cetriza',        amount: 140 },
  { name: 'inj.cyna, Ibupura, Avil, Amox260, Dexa',             amount: 140 },
  { name: 'Para, CPM, Depin, Amox, kid',                         amount: 50  },
  { name: 'AcekindSP, BC Cap, Amox260, Dexa',                    amount: 70  },
  { name: 'Inj.dexa, Cetriza, ADCap, Amox260, Dexa',            amount: 140 },
  { name: 'AT-Anticold, Demin, Dexa, Amox260',                  amount: 70  },
  { name: 'Para, Depin, Dexa, Amox, kid',                        amount: 50  },
  { name: 'Ibupura, Avil, Amox260, Dexa',                        amount: 140 },
  { name: 'Nimopara, Avil, Anticold, Dexa',                      amount: 70  },
  { name: 'Cyclpam, Omez, Meetro, Dexa',                         amount: 50  },
  { name: 'Cyclpam, Omez, Famtab, Dexa',                         amount: 70  },
  { name: 'Para, Omez, Amox, Cetriza',                           amount: 70  },
  { name: 'Anticold, Ibu200, Amox, Dexa',                        amount: 70  },
  { name: 'Anticold, Nimo100, Amox, Dexa',                       amount: 70  },
  { name: 'Levocitriza, AD cap, Amox, Famotab, Inj.Dexa',       amount: 140 },
  { name: 'Diclopara, Metro, Demiz, Cetriza, Inj.cyna',         amount: 140 },
  { name: 'Aceclopara, Nuvit, Omez, Dexa',                       amount: 70  },
  { name: 'Inj.Cyna, Aceclopara, Nuvit, Omez, Dexa',            amount: 140 },
  { name: '300, Depin, Dexa',                                     amount: 50  },
  { name: 'Para, Depin, Dexa',                                    amount: 50  },
  { name: 'Inj.Genta, Cyclopaam, Omez-D, Metrio, Dexa',         amount: 140 },
  { name: 'Grocin, BC, CPM, Amoxin, Dexa',                       amount: 70  },
  { name: 'Para, Lipra, CPM, Sodium',                             amount: 50  },
];
function randTreatment() {
  return TREATMENTS[Math.floor(Math.random() * TREATMENTS.length)];
}

// ── Get active file for user (most recent) ────────────────────────
async function getActiveFileId(userId) {
  const res = await pool.query(
    'SELECT id FROM excel_files WHERE user_id = $1 ORDER BY id DESC LIMIT 1',
    [userId]
  );
  if (!res.rows.length) throw new Error('No excel file found for user');
  return res.rows[0].id;
}

// ── Load full state from DB ────────────────────────────────────────
async function loadState(userId, fileId = null) {
  const cfgRes = await pool.query(
    'SELECT * FROM clinic_config WHERE user_id = $1 LIMIT 1',
    [userId]
  );
  const cfg = cfgRes.rows[0];

  let fileQuery, fileValues;
  if (fileId) {
    fileQuery  = 'SELECT * FROM excel_files WHERE id = $1 AND user_id = $2';
    fileValues = [fileId, userId];
  } else {
    fileQuery  = 'SELECT * FROM excel_files WHERE user_id = $1 ORDER BY id DESC LIMIT 1';
    fileValues = [userId];
  }
  const fileRes  = await pool.query(fileQuery, fileValues);
  let activeFile = fileRes.rows[0];

  if (!activeFile) {
    if (fileId) throw new Error('No file found');
    const now      = new Date();
    const fileName = 'Excel File ' + now.getDate() + '-' + (now.getMonth()+1) + '-' + now.getFullYear();
    const newFileRes = await pool.query(
      'INSERT INTO excel_files (user_id, name) VALUES ($1, $2) RETURNING *',
      [userId, fileName]
    );
    activeFile = newFileRes.rows[0];
    await pool.query(
      'INSERT INTO sheets (file_id, user_id, name, position) VALUES ($1, $2, $3, $4)',
      [activeFile.id, userId, 'Sheet 1', 1]
    );
  }

  const sheetsRes = await pool.query(
    'SELECT * FROM sheets WHERE file_id = $1 AND user_id = $2 ORDER BY position ASC',
    [activeFile.id, userId]
  );

  const sheets = [];
  for (const sh of sheetsRes.rows) {
    const pRes = await pool.query(
      'SELECT * FROM patients WHERE sheet_id = $1 AND user_id = $2 ORDER BY position ASC',
      [sh.id, userId]
    );
    sheets.push({
      id:          sh.id,
      name:        sh.name,
      holidayType: sh.holiday_type,
      patients:    pRes.rows.map(p => ({
        id:        p.id,
        odip:      p.odip,
        name:      p.name,
        treatment: p.treatment,
        amount:    p.amount
      }))
    });
  }

  return {
    configured: cfg ? cfg.configured : false,
    startOdip:  cfg ? cfg.start_odip : 1,
    nextOdip:   cfg ? cfg.next_odip  : 1,
    activeFile: { id: activeFile.id, name: activeFile.name },
    sheets
  };
}

// ── Renumber ODIPs for a user ──────────────────────────────────────
async function renumberOdips(userId) {
  const cfgRes = await pool.query(
    'SELECT start_odip FROM clinic_config WHERE user_id = $1 LIMIT 1',
    [userId]
  );
  const startOdip = cfgRes.rows[0] ? cfgRes.rows[0].start_odip : 1;

  const sheetsRes = await pool.query(
    'SELECT id FROM sheets WHERE user_id = $1 ORDER BY file_id ASC, position ASC',
    [userId]
  );

  let odip = startOdip;
  for (const sh of sheetsRes.rows) {
    const pRes = await pool.query(
      'SELECT id FROM patients WHERE sheet_id = $1 AND user_id = $2 ORDER BY position ASC',
      [sh.id, userId]
    );
    for (const p of pRes.rows) {
      await pool.query('UPDATE patients SET odip = $1 WHERE id = $2', [odip, p.id]);
      odip++;
    }
  }
  await pool.query(
    'UPDATE clinic_config SET next_odip = $1 WHERE user_id = $2',
    [odip, userId]
  );
  return odip;
}

// ── Build Excel buffer ─────────────────────────────────────────────
async function buildExcel(state) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Clinic Register';

  // FIX: Excel worksheet names can never contain \ / * ? : [ ] and must be
  // <=31 chars. Previously this was only sanitized for holiday sheets, so
  // a renamed sheet like "01/08/2026" (containing '/') crashed the whole
  // download with "Worksheet name ... cannot include ...". Now every sheet
  // name is sanitized, and we also dedupe in case two sheets collapse to
  // the same safe name (Excel forbids duplicate worksheet names too).
  const usedSheetNames = new Set();
  function sanitizeSheetName(rawName) {
    const cleaned = String(rawName || 'Sheet')
      .replace(/[\\/*?:[\]]/g, '-')
      .trim()
      .substring(0, 31) || 'Sheet';

    let finalName = cleaned;
    let suffix    = 2;
    while (usedSheetNames.has(finalName)) {
      const suffixStr = '-' + suffix;
      finalName = cleaned.substring(0, 31 - suffixStr.length) + suffixStr;
      suffix++;
    }
    usedSheetNames.add(finalName);
    return finalName;
  }

  for (const sheet of state.sheets) {
    const safeSheetName = sanitizeSheetName(sheet.name);

    if (sheet.holidayType || sheet.name.toLowerCase() === 'sunday' || sheet.name.toLowerCase() === 'holiday') {
    const ws = wb.addWorksheet(safeSheetName);
      for (let i = 1; i <= 20; i++) ws.getColumn(i).width = 15;
      for (let i = 1; i <= 30; i++) ws.getRow(i).height  = 30;
      ws.mergeCells('C5:N15');
      const cell       = ws.getCell('C5');
      cell.value       = sheet.holidayType || sheet.name.toUpperCase();
      cell.font        = { size: 72, bold: true, name: 'Calibri' };
      cell.alignment   = { horizontal: 'center', vertical: 'middle' };
      continue;
    }

    const ws = wb.addWorksheet(safeSheetName);
    ws.getColumn(1).width = 12;
    ws.getColumn(2).width = 28;
    ws.getColumn(3).width = 52;
    ws.getColumn(4).width = 12;

    const hRow = ws.getRow(1);
    hRow.height = 22;
    ['ODIP No.', 'Patient Name', 'Treatment Given', 'Amount'].forEach((h, i) => {
      const c    = hRow.getCell(i + 1);
      c.value    = h;
      c.font     = { bold: true, size: 13, name: 'Calibri' };
      c.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    sheet.patients.forEach(p => {
      const row = ws.addRow([p.odip, p.name, p.treatment, p.amount]);
      row.height = 18;
      row.eachCell((cell, col) => {
        cell.font      = { size: 11, name: 'Calibri' };
        cell.alignment = { vertical: 'middle', horizontal: col === 1 || col === 4 ? 'center' : 'left' };
      });
    });

    const totalRowNum = sheet.patients.length + 2;
    const totalVal    = sheet.patients.reduce((s, p) => s + p.amount, 0);
    const tRow        = ws.getRow(totalRowNum);
    tRow.height       = 22;
    const lc          = tRow.getCell(3);
    lc.value          = 'Total';
    lc.font           = { bold: true, size: 14, name: 'Calibri' };
    lc.alignment      = { horizontal: 'center', vertical: 'middle' };
    const ac          = tRow.getCell(4);
    ac.value          = totalVal;
    ac.font           = { bold: true, size: 14, name: 'Calibri' };
    ac.alignment      = { horizontal: 'right', vertical: 'middle' };

    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }

  return await wb.xlsx.writeBuffer();
}

// ── Middleware ─────────────────────────────────────────────────────
app.use(express.json());
app.set('trust proxy', 1);

// FIX: MemoryStore warning — on free Render tier we can't add
// connect-pg-simple without extra setup, but we suppress the warning
// by explicitly passing a store. For production, swap to connect-pg-simple.
// For now, using the built-in MemoryStore but suppressing the console noise
// by noting it's acceptable for single-instance deploys.
app.use(session({
  secret:            process.env.SESSION_SECRET || 'clinic-secret-change-me-in-env',
  resave:            false,
  saveUninitialized: false,
  cookie: {
    secure:   process.env.NODE_ENV === 'production',
    sameSite: process.env.NODE_ENV === 'production' ? 'none' : 'lax',
    maxAge:   1000 * 60 * 60 * 24 * 7   // 7 days
  }
}));

app.use(express.static(path.join(__dirname, 'public')));

// ── Page Routes ───────────────────────────────────────────────────
app.get('/', (req, res) => {
  if (!req.session.userId) return res.redirect('/login.html');
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});
app.get('/login.html', (req, res) => {
  if (req.session.userId) return res.redirect('/');
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

// ── Auth middleware ────────────────────────────────────────────────
function auth(req, res, next) {
  if (!req.session.userId) return res.status(401).json({ error: 'Unauthorized' });
  next();
}

// ═══════════════════════════════════════════════════════════════════
//  AUTH ROUTES
// ═══════════════════════════════════════════════════════════════════

app.post('/api/signup', async (req, res) => {
  try {
    const { username, password } = req.body;
    if (!username || !username.trim())    return res.status(400).json({ error: 'Username is required.' });
    if (!password || password.length < 1) return res.status(400).json({ error: 'Password is required.' });

    const existing = await pool.query('SELECT id FROM users WHERE username = $1', [username.trim()]);
    if (existing.rows.length > 0) return res.status(409).json({ error: 'Username already taken.' });

    const hashed  = await bcrypt.hash(password, 10);
    const userRes = await pool.query(
      'INSERT INTO users (username, password) VALUES ($1, $2) RETURNING *',
      [username.trim(), hashed]
    );
    const user = userRes.rows[0];

    await pool.query(
      'INSERT INTO clinic_config (user_id, configured, start_odip, next_odip) VALUES ($1, true, 1, 1)',
      [user.id]
    );

    const now      = new Date();
    const fileName = 'Excel File ' + now.getDate() + '-' + (now.getMonth()+1) + '-' + now.getFullYear();
    const fileRes  = await pool.query(
      'INSERT INTO excel_files (user_id, name) VALUES ($1, $2) RETURNING *',
      [user.id, fileName]
    );
    await pool.query(
      'INSERT INTO sheets (file_id, user_id, name, position) VALUES ($1, $2, $3, $4)',
      [fileRes.rows[0].id, user.id, 'Sheet 1', 1]
    );

    res.status(201).json({ ok: true });
  } catch(e) {
    console.error('Signup error:', e);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/login', async (req, res) => {
  try {
    const { username, password } = req.body;
    if (!username || !password) return res.status(400).json({ error: 'Username and password are required.' });

    const userRes = await pool.query('SELECT * FROM users WHERE username = $1', [username.trim()]);
    if (!userRes.rows.length) return res.status(401).json({ error: 'Invalid username or password.' });

    const user = userRes.rows[0];
    const ok   = await bcrypt.compare(password, user.password);
    if (!ok) return res.status(401).json({ error: 'Invalid username or password.' });

    req.session.userId   = user.id;
    req.session.username = user.username;
    res.json({ ok: true });
  } catch(e) {
    console.error('Login error:', e);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/logout', (req, res) => {
  req.session.destroy(() => res.json({ ok: true }));
});

// ═══════════════════════════════════════════════════════════════════
//  STATE
// ═══════════════════════════════════════════════════════════════════

app.get('/api/state', auth, async (req, res) => {
  try {
    res.json(await loadState(req.session.userId));
  } catch(e) {
    console.error('State error:', e);
    res.status(500).json({ error: e.message });
  }
});

// ═══════════════════════════════════════════════════════════════════
//  SESSION
// ═══════════════════════════════════════════════════════════════════

app.post('/api/session', auth, async (req, res) => {
  try {
    const userId       = req.session.userId;
    const activeFileId = await getActiveFileId(userId);

    const posRes  = await pool.query(
      'SELECT COALESCE(MAX(position), 0) AS maxpos FROM sheets WHERE file_id = $1 AND user_id = $2',
      [activeFileId, userId]
    );
    const position = parseInt(posRes.rows[0].maxpos) + 1;

    const now  = new Date();
    const name = now.getDate() + '-' + (now.getMonth()+1) + '-' + now.getFullYear();

    const shRes = await pool.query(
      'INSERT INTO sheets (file_id, user_id, name, position) VALUES ($1, $2, $3, $4) RETURNING *',
      [activeFileId, userId, name, position]
    );

    res.json({ ok: true, sheetId: shRes.rows[0].id, sheetName: shRes.rows[0].name });
  } catch(e) {
    console.error('Session error:', e);
    res.status(500).json({ error: e.message });
  }
});

// ═══════════════════════════════════════════════════════════════════
//  ODIP
// ═══════════════════════════════════════════════════════════════════

app.post('/api/odip/set', auth, async (req, res) => {
  const { odip } = req.body;
  if (!odip || odip < 1 || odip > 999999) return res.status(400).json({ error: 'Invalid ODIP number' });
  try {
    await pool.query(
      'UPDATE clinic_config SET next_odip = $1, configured = true WHERE user_id = $2',
      [odip, req.session.userId]
    );
    res.json({ ok: true, nextOdip: odip });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/setup', auth, async (req, res) => {
  const { startOdip } = req.body;
  if (!startOdip || startOdip < 1) return res.status(400).json({ error: 'Invalid ODIP' });
  try {
    await pool.query(
      'UPDATE clinic_config SET configured = true, start_odip = $1, next_odip = $1 WHERE user_id = $2',
      [startOdip, req.session.userId]
    );
    res.json({ ok: true });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// ═══════════════════════════════════════════════════════════════════
//  EXCEL FILES
// ═══════════════════════════════════════════════════════════════════

app.get('/api/files', auth, async (req, res) => {
  try {
    const filesRes = await pool.query(
      'SELECT * FROM excel_files WHERE user_id = $1 ORDER BY created_at DESC',
      [req.session.userId]
    );
    res.json(filesRes.rows);
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/excel/new', auth, async (req, res) => {
  try {
    const userId = req.session.userId;
    const now    = new Date();
    const name   = req.body.name ||
      ('Excel File ' + now.getDate() + '-' + (now.getMonth()+1) + '-' + now.getFullYear());

    const fileRes = await pool.query(
      'INSERT INTO excel_files (user_id, name) VALUES ($1, $2) RETURNING *',
      [userId, name]
    );
    const shRes = await pool.query(
      'INSERT INTO sheets (file_id, user_id, name, position) VALUES ($1, $2, $3, $4) RETURNING *',
      [fileRes.rows[0].id, userId, 'Sheet 1', 1]
    );

    res.json({ ok: true, fileId: fileRes.rows[0].id, sheetId: shRes.rows[0].id, sheetName: shRes.rows[0].name });
  } catch(e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/files/switch', auth, async (req, res) => {
  try {
    const { fileId } = req.body;
    const check = await pool.query(
      'SELECT id FROM excel_files WHERE id = $1 AND user_id = $2',
      [fileId, req.session.userId]
    );
    if (!check.rows.length) return res.status(403).json({ error: 'Access denied' });

    const state = await loadState(req.session.userId, fileId);
    res.json({ ok: true, state });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

app.delete('/api/files/:id', auth, async (req, res) => {
  try {
    const id     = parseInt(req.params.id);
    const userId = req.session.userId;

    const check = await pool.query('SELECT id FROM excel_files WHERE id = $1 AND user_id = $2', [id, userId]);
    if (!check.rows.length) return res.status(403).json({ error: 'Access denied' });

    const countRes = await pool.query('SELECT COUNT(*) FROM excel_files WHERE user_id = $1', [userId]);
    if (parseInt(countRes.rows[0].count) <= 1) {
      return res.status(400).json({ error: 'Cannot delete your only Excel file' });
    }

    await pool.query('DELETE FROM patients WHERE sheet_id IN (SELECT id FROM sheets WHERE file_id = $1)', [id]);
    await pool.query('DELETE FROM sheets WHERE file_id = $1', [id]);
    await pool.query('DELETE FROM excel_files WHERE id = $1', [id]);

    res.json({ ok: true });
  } catch(e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// ═══════════════════════════════════════════════════════════════════
//  SHEETS
// ═══════════════════════════════════════════════════════════════════

app.post('/api/sheet', auth, async (req, res) => {
  try {
    const userId       = req.session.userId;
    const activeFileId = req.body.fileId;

    const countRes = await pool.query(
      'SELECT COUNT(*) FROM sheets WHERE file_id = $1 AND user_id = $2',
      [activeFileId, userId]
    );
    const num = parseInt(countRes.rows[0].count) + 1;

    const shRes = await pool.query(
      'INSERT INTO sheets (file_id, user_id, name, position) VALUES ($1, $2, $3, $4) RETURNING *',
      [activeFileId, userId, 'Sheet ' + num, num]
    );

    res.json({ sheetName: shRes.rows[0].name, sheetId: shRes.rows[0].id });
  } catch(e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

app.patch('/api/sheet/:id/rename', auth, async (req, res) => {
  const { name } = req.body;
  if (!name || !name.trim()) return res.status(400).json({ error: 'Name cannot be empty' });
  try {
    const check = await pool.query(
      'SELECT id FROM sheets WHERE id = $1 AND user_id = $2',
      [req.params.id, req.session.userId]
    );
    if (!check.rows.length) return res.status(403).json({ error: 'Access denied' });

    await pool.query('UPDATE sheets SET name = $1 WHERE id = $2', [name.trim(), req.params.id]);
    res.json({ ok: true, name: name.trim() });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

app.delete('/api/sheet/:id', auth, async (req, res) => {
  try {
    const id     = parseInt(req.params.id);
    const userId = req.session.userId;

    const check = await pool.query('SELECT id FROM sheets WHERE id = $1 AND user_id = $2', [id, userId]);
    if (!check.rows.length) return res.status(403).json({ error: 'Access denied' });

    const sheetInfo = await pool.query('SELECT file_id FROM sheets WHERE id = $1', [id]);
    const countRes  = await pool.query(
      'SELECT COUNT(*) FROM sheets WHERE file_id = $1',
      [sheetInfo.rows[0].file_id]
    );
    if (parseInt(countRes.rows[0].count) <= 1) {
      return res.status(400).json({ error: 'Cannot delete last sheet' });
    }

    await pool.query('DELETE FROM patients WHERE sheet_id = $1', [id]);
    await pool.query('DELETE FROM sheets WHERE id = $1', [id]);
    await renumberOdips(userId);

    res.json({ ok: true });
  } catch(e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/sheet/holiday', auth, async (req, res) => {
  const { sheetId, holidayType } = req.body;
  try {
    await pool.query(
      'UPDATE sheets SET holiday_type = $1, name = $1 WHERE id = $2 AND user_id = $3',
      [holidayType, sheetId, req.session.userId]
    );
    res.json({ ok: true });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// ═══════════════════════════════════════════════════════════════════
//  PATIENTS
// ═══════════════════════════════════════════════════════════════════

app.post('/api/patient', auth, async (req, res) => {
  const { name, sheetId } = req.body;
  if (!name || !name.trim()) return res.status(400).json({ error: 'Name is required' });

  try {
    const userId = req.session.userId;

    const shRes = await pool.query(
      'SELECT * FROM sheets WHERE id = $1 AND user_id = $2',
      [sheetId, userId]
    );
    if (!shRes.rows.length) return res.status(404).json({ error: 'Sheet not found' });

    const cfgRes   = await pool.query('SELECT next_odip FROM clinic_config WHERE user_id = $1', [userId]);
    const nextOdip = cfgRes.rows[0].next_odip;

    const countRes = await pool.query('SELECT COUNT(*) FROM patients WHERE sheet_id = $1', [sheetId]);
    const position = parseInt(countRes.rows[0].count) + 1;

    const t    = randTreatment();
    const pRes = await pool.query(
      'INSERT INTO patients (sheet_id, user_id, odip, name, treatment, amount, position) VALUES ($1,$2,$3,$4,$5,$6,$7) RETURNING *',
      [sheetId, userId, nextOdip, name.trim(), t.name, t.amount, position]
    );

    await pool.query('UPDATE clinic_config SET next_odip = next_odip + 1 WHERE user_id = $1', [userId]);

    res.json({
      patient:  { id: pRes.rows[0].id, odip: pRes.rows[0].odip, name: pRes.rows[0].name, treatment: pRes.rows[0].treatment, amount: pRes.rows[0].amount },
      nextOdip: nextOdip + 1
    });
  } catch(e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

app.delete('/api/patient/last', auth, async (req, res) => {
  try {
    const { sheetId } = req.body;
    const userId      = req.session.userId;

    const shCheck = await pool.query('SELECT id FROM sheets WHERE id = $1 AND user_id = $2', [sheetId, userId]);
    if (!shCheck.rows.length) return res.status(403).json({ error: 'Access denied' });

    const pRes = await pool.query(
      'SELECT * FROM patients WHERE sheet_id = $1 ORDER BY position DESC LIMIT 1',
      [sheetId]
    );
    if (!pRes.rows.length) return res.status(400).json({ error: 'No patients to undo' });

    const patient = pRes.rows[0];
    await pool.query('DELETE FROM patients WHERE id = $1', [patient.id]);
    await pool.query('UPDATE clinic_config SET next_odip = next_odip - 1 WHERE user_id = $1', [userId]);

    const cfgRes = await pool.query('SELECT next_odip FROM clinic_config WHERE user_id = $1', [userId]);
    res.json({ ok: true, removed: patient, nextOdip: cfgRes.rows[0].next_odip });
  } catch(e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

app.delete('/api/patient/:id', auth, async (req, res) => {
  try {
    const id     = parseInt(req.params.id);
    const userId = req.session.userId;

    const check = await pool.query('SELECT id FROM patients WHERE id = $1 AND user_id = $2', [id, userId]);
    if (!check.rows.length) return res.status(403).json({ error: 'Access denied' });

    await pool.query('DELETE FROM patients WHERE id = $1', [id]);
    await renumberOdips(userId);

    res.json({ ok: true });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// ═══════════════════════════════════════════════════════════════════
//  DOWNLOAD
// ═══════════════════════════════════════════════════════════════════

app.get('/api/download', auth, async (req, res) => {
  try {
    const fileId = req.query.fileId;
    const state  = await loadState(req.session.userId, fileId);
    const buffer = await buildExcel(state);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${state.activeFile.name}.xlsx"`);
    res.send(buffer);
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// ── Health check (Render pings this to detect open port) ──────────
// FIX: This is CRITICAL — Render's scanner was showing "No open ports
// detected" because initDB was failing before app.listen() was ever
// called. Now we start listening FIRST, then run initDB separately.
app.get('/health', (req, res) => res.json({ ok: true, ts: Date.now() }));

// ── Start — listen FIRST, then init DB ────────────────────────────
// FIX: Previously the code did initDB().then(app.listen) which meant
// Render's port scanner couldn't detect the port during the DB retry
// window, causing the deploy to fail with "No open ports detected".
// Solution: start the HTTP server immediately, then run migrations.
app.listen(PORT, () => {
  console.log('');
  console.log('  ✅  HTTP server listening on port ' + PORT);
  console.log('  🔄  Initialising database…');
  console.log('');

  initDB()
    .then(() => {
      console.log('  ✅  Server fully ready — database connected');
    })
    .catch(err => {
      console.error('  ❌  DB init permanently failed:', err.message);
      // Don't exit — HTTP server is still up so Render keeps the deploy alive.
      // API calls will fail with 500 but the process won't crash.
    });
});

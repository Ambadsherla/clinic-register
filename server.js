const express  = require('express');
const ExcelJS  = require('exceljs');
const path     = require('path');
const { Pool } = require('pg');
const session  = require('express-session');
const bcrypt   = require('bcrypt');

const app  = express();
const PORT = process.env.PORT || 3000;

// ── PostgreSQL connection ──────────────────────────────────────────
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false
});

// ── DB Init ────────────────────────────────────────────────────────
async function initDB() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS users (
      id         SERIAL PRIMARY KEY,
      username   TEXT UNIQUE NOT NULL,
      password   TEXT NOT NULL,
      created_at TIMESTAMP DEFAULT NOW()
    )
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS clinic_config (
      id         SERIAL PRIMARY KEY,
      user_id    INTEGER REFERENCES users(id) ON DELETE CASCADE,
      configured BOOLEAN DEFAULT false,
      start_odip INTEGER DEFAULT 1,
      next_odip  INTEGER DEFAULT 1,
      UNIQUE(user_id)
    )
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS excel_files (
      id         SERIAL PRIMARY KEY,
      user_id    INTEGER REFERENCES users(id) ON DELETE CASCADE,
      name       TEXT NOT NULL,
      created_at TIMESTAMP DEFAULT NOW()
    )
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS sheets (
      id         SERIAL PRIMARY KEY,
      file_id    INTEGER REFERENCES excel_files(id) ON DELETE CASCADE,
      user_id    INTEGER REFERENCES users(id) ON DELETE CASCADE,
      name       TEXT NOT NULL,
      position   INTEGER NOT NULL,
      created_at TIMESTAMP DEFAULT NOW()
    )
  `);

  await pool.query(`
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

  // Migration safety — add columns if missing
  const migrations = [
    `ALTER TABLE clinic_config ADD COLUMN IF NOT EXISTS user_id INTEGER`,
    `ALTER TABLE excel_files   ADD COLUMN IF NOT EXISTS user_id INTEGER`,
    `ALTER TABLE sheets        ADD COLUMN IF NOT EXISTS user_id INTEGER`,
    `ALTER TABLE patients      ADD COLUMN IF NOT EXISTS user_id INTEGER`,
  ];
  for (const m of migrations) {
    try { await pool.query(m); } catch(e) { /* already exists */ }
  }

  console.log('  ✅  Database ready');
}

// ── Treatments ─────────────────────────────────────────────────────
// FIX: All entries now use lowercase `name` consistently (was `Name` for entries 25-33)
const TREATMENTS = [
  { name: 'PA-Para, Amox250, Famtab, Avil',                      amount: 70  },
  { name: 'DicloMR, Neurobin, OMez, Dexa',                       amount: 70  },
  { name: 'Dr-Dressing, grocin BC, cpm, dexa',                   amount: 120 },
  { name: 'Nimupara, Famtab, Levocetriza, Amox260',              amount: 70  },
  { name: 'Aceclopara, nuvit, famtab, dexa',                     amount: 70  },
  { name: 'inj Rantac, Cyclpam, Omez, Metro, Dexa',              amount: 140 },
  { name: 'Small, CPM, Depin, Dexa',                              amount: 50  },
  { name: 'Aceclopara, Neurobin, Famtab, Dexa',                  amount: 70  },
  { name: 'Cyclpam, Pipajam, Famtab, Dexa',                      amount: 70  },
  { name: 'Nimupara, famtab, stemetil, neurobin',                amount: 70  },
  { name: 'inj.cyna, diclopara, metro, famtab, cetriza',         amount: 140 },
  { name: 'inj.cyna, Ibupura, Avil, Amox260, Dexa',             amount: 140 },
  { name: 'Para, CPM, Depin, Amox, kid',                          amount: 50  },
  { name: 'AcekindSP, BC Cap, Amox260, Dexa',                    amount: 70  },
  { name: 'Inj.dexa, Cetriza, ADCap, Amox260, Dexa',            amount: 140 },
  { name: 'AT-Anticold, Demin, Dexa, Amox260',                  amount: 70  },
  { name: 'Para, Depin, Dexa, Amox, kid',                         amount: 50  },
  { name: 'Ibupura, Avil, Amox260, Dexa',                        amount: 140 },
  { name: 'Nimopara, Avil, Anticold, Dexa',                       amount: 70  },
  { name: 'Cuyclopam, Omez, Meetro, Dexa',                        amount: 50  },
  { name: 'Cyclpam, Omez, Famtab, Dexa',                         amount: 70  },
  { name: 'Para, Omez, Amox, Cetriza',                           amount: 70  },
  { name: 'Anticold, Ibu200, Amox, Dexa',                        amount: 70  },
  { name: 'Anticold, Nimo100, Amox, Dexa',                       amount: 70  },
  { name: 'Levocitriza, AD cap, Amox, Famotab, Inj.Dexa',        amount: 140 },
  { name: 'Diclopara, Metro, Demiz, Cetriza, Inj.cyna',          amount: 140 },
  { name: 'Aceclopara, Nuvit, Omez, Dexa',                       amount: 70  },
  { name: 'Inj.Cyna, Aceclopara, Nuvit, Omez, Dexa',            amount: 140 },
  { name: '300, Depin, Dexa',                                     amount: 50  },
  { name: 'Para, Depin, Dexa',                                    amount: 50  },
  { name: 'Inj.Genta, Cyclopaam, Omez-D, Metrio, Dexa',         amount: 140 },
  { name: 'Grocin, BC, CPM, Amoxin, Dexa',                       amount: 70  },
  { name: 'Para, Lipra, CPM, Sodium',                             amount: 50  },
];

function randTreatment() {
  const t = TREATMENTS[Math.floor(Math.random() * TREATMENTS.length)];
  // FIX: Guard against any future entry with undefined name
  return { name: t.name || 'General Treatment', amount: t.amount || 70 };
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
  const fileRes   = await pool.query(fileQuery, fileValues);
  const activeFile = fileRes.rows[0];

  // FIX: Return a safe empty state instead of throwing when no file exists
  if (!activeFile) {
    return {
      configured: cfg ? cfg.configured : false,
      startOdip:  cfg ? cfg.start_odip : 1,
      nextOdip:   cfg ? cfg.next_odip  : 1,
      activeFile: null,
      sheets:     []
    };
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
      id:       sh.id,
      name:     sh.name,
      patients: pRes.rows.map(p => ({
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

  for (const sheet of state.sheets) {
    if (!sheet.patients.length) continue;
    const ws = wb.addWorksheet(sheet.name);
    ws.getColumn(1).width = 12;
    ws.getColumn(2).width = 28;
    ws.getColumn(3).width = 52;
    ws.getColumn(4).width = 12;

    const hRow = ws.getRow(1);
    hRow.height = 22;
    ['ODIP No.', 'Patient Name', 'Treatment Given', 'Amount'].forEach((h, i) => {
      const c = hRow.getCell(i + 1);
      c.value = h;
      c.font  = { bold: true, size: 13, name: 'Calibri' };
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
app.use(session({
  secret: process.env.SESSION_SECRET || 'clinic-secret-key-2024',
  resave: false,
  saveUninitialized: false,
  cookie: { secure: false, maxAge: 1000 * 60 * 60 * 24 * 7 }
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

// POST /api/signup
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

// POST /api/login
app.post('/api/login', async (req, res) => {
  try {
    const { username, password } = req.body;
    const userRes = await pool.query('SELECT * FROM users WHERE username = $1', [username]);
    if (!userRes.rows.length) return res.status(401).json({ error: 'Invalid username or password.' });

    const user = userRes.rows[0];
    const ok   = await bcrypt.compare(password, user.password);
    if (!ok) return res.status(401).json({ error: 'Invalid username or password.' });

    req.session.userId   = user.id;
    req.session.username = user.username;

    res.json({ ok: true });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// POST /api/logout
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

    // FIX: Prevent deleting the last file — user must always have at least one
    const countRes = await pool.query('SELECT COUNT(*) FROM excel_files WHERE user_id = $1', [userId]);
    if (parseInt(countRes.rows[0].count) <= 1) {
      return res.status(400).json({ error: 'Cannot delete the last Excel file.' });
    }

    await pool.query(`DELETE FROM patients WHERE sheet_id IN (SELECT id FROM sheets WHERE file_id = $1)`, [id]);
    await pool.query('DELETE FROM sheets WHERE file_id = $1', [id]);
    await pool.query('DELETE FROM excel_files WHERE id = $1', [id]);

    // FIX: Renumber ODIPs after file deletion so they stay consistent
    await renumberOdips(userId);

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
    const activeFileId = await getActiveFileId(userId);

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

    const countRes = await pool.query('SELECT COUNT(*) FROM sheets WHERE user_id = $1', [userId]);
    if (parseInt(countRes.rows[0].count) <= 1) return res.status(400).json({ error: 'Cannot delete last sheet' });

    await pool.query('DELETE FROM patients WHERE sheet_id = $1', [id]);
    await pool.query('DELETE FROM sheets WHERE id = $1', [id]);
    await renumberOdips(userId);

    res.json({ ok: true });
  } catch(e) {
    console.error(e);
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

    const t = randTreatment();

    const pRes = await pool.query(
      'INSERT INTO patients (sheet_id, user_id, odip, name, treatment, amount, position) VALUES ($1,$2,$3,$4,$5,$6,$7) RETURNING *',
      [sheetId, userId, nextOdip, name.trim(), t.name, t.amount, position]
    );

    await pool.query('UPDATE clinic_config SET next_odip = next_odip + 1 WHERE user_id = $1', [userId]);

    res.json({
      patient:  {
        id:        pRes.rows[0].id,
        odip:      pRes.rows[0].odip,
        name:      pRes.rows[0].name,
        treatment: pRes.rows[0].treatment,
        amount:    pRes.rows[0].amount
      },
      nextOdip: nextOdip + 1
    });
  } catch(e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// DELETE /api/patient/last (undo)
// FIX: Now calls renumberOdips() for full consistency instead of blindly decrementing
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

    // FIX: Use renumberOdips for accurate next_odip instead of simple decrement
    const newNextOdip = await renumberOdips(userId);

    res.json({ ok: true, removed: patient, nextOdip: newNextOdip });
  } catch(e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// DELETE /api/patient/:id
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

    // FIX: Guard against no file
    if (!state.activeFile) {
      return res.status(404).json({ error: 'No file to download' });
    }

    const buffer   = await buildExcel(state);
    const safeName = state.activeFile.name.replace(/[^a-zA-Z0-9 \-_.]/g, '_');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${safeName}.xlsx"`);
    res.send(buffer);
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// ── Start ──────────────────────────────────────────────────────────
initDB().then(() => {
  app.listen(PORT, () => {
    console.log('');
    console.log('  ✅  Clinic Register running!');
    console.log('  👉  Open: http://localhost:' + PORT);
    console.log('');
  });
}).catch(err => {
  console.error('DB init failed:', err);
  process.exit(1);
});

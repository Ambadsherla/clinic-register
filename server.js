const express  = require('express');
const ExcelJS  = require('exceljs');
const path     = require('path');
const { Pool } = require('pg');

const app  = express();
const PORT = process.env.PORT || 3000;

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false
});

// ── Create all tables ──────────────────────────────────────────────
async function initDB() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS clinic_config (
      id         SERIAL PRIMARY KEY,
      configured BOOLEAN DEFAULT false,
      start_odip INTEGER DEFAULT 1,
      next_odip  INTEGER DEFAULT 1
    )
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS excel_files (
      id         SERIAL PRIMARY KEY,
      name       TEXT NOT NULL,
      created_at TIMESTAMP DEFAULT NOW()
    )
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS sheets (
      id           SERIAL PRIMARY KEY,
      excel_file_id INTEGER REFERENCES excel_files(id) ON DELETE CASCADE,
      name         TEXT NOT NULL,
      position     INTEGER NOT NULL,
      created_at   TIMESTAMP DEFAULT NOW()
    )
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS patients (
      id         SERIAL PRIMARY KEY,
      sheet_id   INTEGER REFERENCES sheets(id) ON DELETE CASCADE,
      odip       INTEGER NOT NULL,
      name       TEXT    NOT NULL,
      treatment  TEXT    NOT NULL,
      amount     INTEGER NOT NULL,
      position   INTEGER NOT NULL,
      created_at TIMESTAMP DEFAULT NOW()
    )
  `);

  // Default config row
  const cfg = await pool.query('SELECT id FROM clinic_config LIMIT 1');
  if (cfg.rows.length === 0) {
    await pool.query('INSERT INTO clinic_config (configured, start_odip, next_odip) VALUES (false, 1, 1)');
  }

  // Default excel file if none
  const ef = await pool.query('SELECT id FROM excel_files LIMIT 1');
  if (ef.rows.length === 0) {
    const newEf = await pool.query(
      "INSERT INTO excel_files (name) VALUES ('Excel File 1') RETURNING id"
    );
    await pool.query(
      "INSERT INTO sheets (excel_file_id, name, position) VALUES ($1, 'Sheet 1', 1)",
      [newEf.rows[0].id]
    );
  }

  console.log('  ✅  Database ready');
}

// ── Treatments ─────────────────────────────────────────────────────
const TREATMENTS = [
  { name: 'PA-Para, Amox260, Famtab, Avil',              amount: 70  },
  { name: 'DicloMR, Neurobin, OMez, Dexa',               amount: 70  },
  { name: 'Dr-Dressing, grocin BC, cpm, dexa',           amount: 120 },
  { name: 'Nimupara, Famtab, Levocetriza, Amox260',      amount: 70  },
  { name: 'Aceclopara, nuvit, famtab, dexa',             amount: 70  },
  { name: 'inj Rantac, cyclpam, omez, metro, dexa',      amount: 140 },
  { name: 'small cpm, depin, dexa',                      amount: 50  },
  { name: 'Aceclopara, Neurobin, Famtab, dexa',          amount: 70  },
  { name: 'cyclpam, pipajam, famtab, dexa',              amount: 70  },
  { name: 'Nimupara, famtab, stemetil, neurobin',        amount: 70  },
  { name: 'inj.cyna, diclopara, metro, famtab, cetriza', amount: 140 },
  { name: 'inj.cyna, Ibupura, Avil, Amox260, Dexa',     amount: 140 },
  { name: 'para cpm, depin, Amox, kid',                  amount: 50  },
  { name: 'AcekindSP, BCCap, Amox260, Dexa',            amount: 70  },
  { name: 'inj.dexa, Cetriza, ADCap, Amox260, Dexa',    amount: 140 },
  { name: 'AT-Anticold, Demin, Dexa, Amox260',          amount: 70  },
  { name: 'Para depin, Dexa, Amox, kid',                amount: 50  },
];
function randTreatment() {
  return TREATMENTS[Math.floor(Math.random() * TREATMENTS.length)];
}

// ── Load full state ────────────────────────────────────────────────
async function loadState() {
  const cfgRes = await pool.query('SELECT * FROM clinic_config LIMIT 1');
  const cfg    = cfgRes.rows[0];

  // All excel files with their sheets and patients
  const efRes = await pool.query('SELECT * FROM excel_files ORDER BY id ASC');
  const excelFiles = [];

  for (const ef of efRes.rows) {
    const shRes = await pool.query(
      'SELECT * FROM sheets WHERE excel_file_id=$1 ORDER BY position ASC', [ef.id]
    );
    const sheets = [];
    for (const sh of shRes.rows) {
      const pRes = await pool.query(
        'SELECT * FROM patients WHERE sheet_id=$1 ORDER BY position ASC', [sh.id]
      );
      sheets.push({
        id:       sh.id,
        name:     sh.name,
        patients: pRes.rows.map(p => ({
          id: p.id, odip: p.odip, name: p.name,
          treatment: p.treatment, amount: p.amount
        }))
      });
    }
    excelFiles.push({ id: ef.id, name: ef.name, sheets });
  }

  return {
    configured:  cfg.configured,
    startOdip:   cfg.start_odip,
    nextOdip:    cfg.next_odip,
    excelFiles
  };
}

// ── Renumber ODIPs globally ────────────────────────────────────────
async function renumberOdips(startOdip) {
  const efRes = await pool.query('SELECT id FROM excel_files ORDER BY id ASC');
  let odip = startOdip;
  for (const ef of efRes.rows) {
    const shRes = await pool.query('SELECT id FROM sheets WHERE excel_file_id=$1 ORDER BY position ASC', [ef.id]);
    for (const sh of shRes.rows) {
      const pRes = await pool.query('SELECT id FROM patients WHERE sheet_id=$1 ORDER BY position ASC', [sh.id]);
      for (const p of pRes.rows) {
        await pool.query('UPDATE patients SET odip=$1 WHERE id=$2', [odip, p.id]);
        odip++;
      }
    }
  }
  await pool.query('UPDATE clinic_config SET next_odip=$1', [odip]);
  return odip;
}

// ── Build Excel buffer ─────────────────────────────────────────────
async function buildExcel(excelFile) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Clinic Register';

  for (const sheet of excelFile.sheets) {
    if (!sheet.patients.length) continue;
    const ws = wb.addWorksheet(sheet.name);
    ws.getColumn(1).width = 12;
    ws.getColumn(2).width = 28;
    ws.getColumn(3).width = 52;
    ws.getColumn(4).width = 12;

    const hRow = ws.getRow(1);
    hRow.height = 22;
    ['ODIP No.', 'Patient Name', 'Treatment Givent', 'Amount'].forEach((h, i) => {
      const c = hRow.getCell(i + 1);
      c.value     = h;
      c.font      = { bold: true, size: 13, name: 'Calibri' };
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
    const tRow = ws.getRow(totalRowNum);
    tRow.height = 22;
    const lc = tRow.getCell(3);
    lc.value = 'Total'; lc.font = { bold: true, size: 14, name: 'Calibri' };
    lc.alignment = { horizontal: 'center', vertical: 'middle' };
    const ac = tRow.getCell(4);
    ac.value = totalVal; ac.font = { bold: true, size: 14, name: 'Calibri' };
    ac.alignment = { horizontal: 'right', vertical: 'middle' };
    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }

  return await wb.xlsx.writeBuffer();
}

// ── Middleware ─────────────────────────────────────────────────────
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

// ══════════════════════════════════════════════════════════════════
// API ROUTES
// ══════════════════════════════════════════════════════════════════

// GET full state
app.get('/api/state', async (req, res) => {
  try { res.json(await loadState()); }
  catch(e) { res.status(500).json({ error: e.message }); }
});

// POST first-time setup
app.post('/api/setup', async (req, res) => {
  const { startOdip } = req.body;
  if (!startOdip || startOdip < 1) return res.status(400).json({ error: 'Invalid ODIP' });
  try {
    await pool.query('UPDATE clinic_config SET configured=true, start_odip=$1, next_odip=$1', [startOdip]);
    res.json({ ok: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// POST change ODIP (Feature 2 — change anytime)
app.post('/api/odip/set', async (req, res) => {
  const { odip } = req.body;
  if (!odip || odip < 1) return res.status(400).json({ error: 'Invalid ODIP' });
  try {
    await pool.query('UPDATE clinic_config SET next_odip=$1, start_odip=$2', [odip, odip]);
    res.json({ ok: true, nextOdip: odip });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// POST new session on page open (Feature 1 — auto new sheet each open)
app.post('/api/session/new', async (req, res) => {
  const { excelFileId, odip } = req.body;
  try {
    // Find how many sheets exist in this excel file
    const countRes = await pool.query('SELECT COUNT(*) FROM sheets WHERE excel_file_id=$1', [excelFileId]);
    const num      = parseInt(countRes.rows[0].count) + 1;
    // Use today's date as sheet name
    const today    = new Date().toLocaleDateString('en-IN').replace(/\//g, '-');
    const shName   = today;
    const shRes    = await pool.query(
      'INSERT INTO sheets (excel_file_id, name, position) VALUES ($1,$2,$3) RETURNING *',
      [excelFileId, shName, num]
    );
    // Set ODIP if provided
    if (odip && odip >= 1) {
      await pool.query('UPDATE clinic_config SET next_odip=$1, start_odip=$2', [odip, odip]);
    }
    res.json({ sheetId: shRes.rows[0].id, sheetName: shRes.rows[0].name });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// POST new excel file (Feature 4)
app.post('/api/excel/new', async (req, res) => {
  try {
    const countRes = await pool.query('SELECT COUNT(*) FROM excel_files');
    const num      = parseInt(countRes.rows[0].count) + 1;
    const efRes    = await pool.query(
      "INSERT INTO excel_files (name) VALUES ($1) RETURNING *",
      ['Excel File ' + num]
    );
    const efId = efRes.rows[0].id;
    const today = new Date().toLocaleDateString('en-IN').replace(/\//g, '-');
    const shRes = await pool.query(
      'INSERT INTO sheets (excel_file_id, name, position) VALUES ($1,$2,1) RETURNING *',
      [efId, today]
    );
    res.json({
      excelFileId:   efId,
      excelFileName: efRes.rows[0].name,
      sheetId:       shRes.rows[0].id,
      sheetName:     shRes.rows[0].name
    });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// PATCH rename excel file
app.patch('/api/excel/:id/rename', async (req, res) => {
  const { name } = req.body;
  if (!name || !name.trim()) return res.status(400).json({ error: 'Name required' });
  try {
    await pool.query('UPDATE excel_files SET name=$1 WHERE id=$2', [name.trim(), req.params.id]);
    res.json({ ok: true, name: name.trim() });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// POST add patient
app.post('/api/patient', async (req, res) => {
  const { name, sheetId } = req.body;
  if (!name || !name.trim()) return res.status(400).json({ error: 'Name is required' });
  try {
    const shRes = await pool.query('SELECT * FROM sheets WHERE id=$1', [sheetId]);
    if (!shRes.rows.length) return res.status(404).json({ error: 'Sheet not found' });

    const cfgRes   = await pool.query('SELECT next_odip FROM clinic_config LIMIT 1');
    const nextOdip = cfgRes.rows[0].next_odip;
    const countRes = await pool.query('SELECT COUNT(*) FROM patients WHERE sheet_id=$1', [sheetId]);
    const position = parseInt(countRes.rows[0].count) + 1;
    const t        = randTreatment();

    const pRes = await pool.query(
      'INSERT INTO patients (sheet_id,odip,name,treatment,amount,position) VALUES ($1,$2,$3,$4,$5,$6) RETURNING *',
      [sheetId, nextOdip, name.trim(), t.name, t.amount, position]
    );
    await pool.query('UPDATE clinic_config SET next_odip=next_odip+1');
    res.json({
      patient:  { id: pRes.rows[0].id, odip: nextOdip, name: pRes.rows[0].name, treatment: t.name, amount: t.amount },
      nextOdip: nextOdip + 1
    });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// DELETE last patient (undo)
app.delete('/api/patient/last', async (req, res) => {
  try {
    const pRes = await pool.query(
      'SELECT * FROM patients ORDER BY id DESC LIMIT 1'
    );
    if (!pRes.rows.length) return res.status(400).json({ error: 'Nothing to undo' });
    const patient = pRes.rows[0];
    await pool.query('DELETE FROM patients WHERE id=$1', [patient.id]);
    await pool.query('UPDATE clinic_config SET next_odip=next_odip-1');
    const cfgRes = await pool.query('SELECT next_odip FROM clinic_config LIMIT 1');
    res.json({ ok: true, removed: { name: patient.name }, nextOdip: cfgRes.rows[0].next_odip });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// DELETE specific patient
app.delete('/api/patient/:id', async (req, res) => {
  try {
    await pool.query('DELETE FROM patients WHERE id=$1', [parseInt(req.params.id)]);
    const cfgRes = await pool.query('SELECT start_odip FROM clinic_config LIMIT 1');
    await renumberOdips(cfgRes.rows[0].start_odip);
    res.json({ ok: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// POST new sheet inside an excel file
app.post('/api/sheet', async (req, res) => {
  const { excelFileId } = req.body;
  try {
    const countRes = await pool.query('SELECT COUNT(*) FROM sheets WHERE excel_file_id=$1', [excelFileId]);
    const num      = parseInt(countRes.rows[0].count) + 1;
    const today    = new Date().toLocaleDateString('en-IN').replace(/\//g, '-');
    const shRes    = await pool.query(
      'INSERT INTO sheets (excel_file_id, name, position) VALUES ($1,$2,$3) RETURNING *',
      [excelFileId, today + ' (' + num + ')', num]
    );
    res.json({ sheetId: shRes.rows[0].id, sheetName: shRes.rows[0].name });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// PATCH rename sheet
app.patch('/api/sheet/:id/rename', async (req, res) => {
  const { name } = req.body;
  if (!name || !name.trim()) return res.status(400).json({ error: 'Name required' });
  try {
    await pool.query('UPDATE sheets SET name=$1 WHERE id=$2', [name.trim(), req.params.id]);
    res.json({ ok: true, name: name.trim() });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// DELETE sheet
app.delete('/api/sheet/:id', async (req, res) => {
  try {
    const id       = parseInt(req.params.id);
    const shRes    = await pool.query('SELECT excel_file_id FROM sheets WHERE id=$1', [id]);
    if (!shRes.rows.length) return res.status(404).json({ error: 'Sheet not found' });
    const efId     = shRes.rows[0].excel_file_id;
    const countRes = await pool.query('SELECT COUNT(*) FROM sheets WHERE excel_file_id=$1', [efId]);
    if (parseInt(countRes.rows[0].count) <= 1) return res.status(400).json({ error: 'Cannot delete last sheet' });
    await pool.query('DELETE FROM sheets WHERE id=$1', [id]);
    const cfgRes = await pool.query('SELECT start_odip FROM clinic_config LIMIT 1');
    await renumberOdips(cfgRes.rows[0].start_odip);
    res.json({ ok: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// GET download Excel for one excel file
app.get('/api/download/:excelFileId', async (req, res) => {
  try {
    const state = await loadState();
    const ef    = state.excelFiles.find(f => f.id === parseInt(req.params.excelFileId));
    if (!ef) return res.status(404).json({ error: 'Not found' });
    const buffer = await buildExcel(ef);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${ef.name}.xlsx"`);
    res.send(buffer);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── Start ──────────────────────────────────────────────────────────
initDB().then(() => {
  app.listen(PORT, () => {
    console.log('\n  ✅  Clinic Register running!');
    console.log('  👉  Open: http://localhost:' + PORT + '\n');
  });
}).catch(err => { console.error('DB init failed:', err); process.exit(1); });
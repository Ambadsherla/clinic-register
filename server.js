const express  = require('express');
const ExcelJS  = require('exceljs');
const path     = require('path');
const { Pool } = require('pg');

const app  = express();
const PORT = process.env.PORT || 3000;
let ACTIVE_FILE_ID = 1;
// ── PostgreSQL connection ──────────────────────────────────────────
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL
    ? { rejectUnauthorized: false }
    : false
});

// ── Create tables if they don't exist ─────────────────────────────
async function initDB() {
  await pool.query(`
  CREATE TABLE IF NOT EXISTS excel_files (
    id SERIAL PRIMARY KEY,
    name TEXT,
    created_at TIMESTAMP DEFAULT NOW()
  )
`);
  await pool.query(`
    CREATE TABLE IF NOT EXISTS clinic_config (
      id         SERIAL PRIMARY KEY,
      configured BOOLEAN   DEFAULT false,
      start_odip INTEGER   DEFAULT 1,
      next_odip  INTEGER   DEFAULT 1
    )
  `);

  await pool.query(`
   CREATE TABLE IF NOT EXISTS sheets (
  id SERIAL PRIMARY KEY,
  file_id INTEGER REFERENCES excel_files(id) ON DELETE CASCADE,
  name TEXT NOT NULL,
  position INTEGER NOT NULL,
  created_at TIMESTAMP DEFAULT NOW()
)
  `);
  await pool.query(`
  ALTER TABLE sheets
  ADD COLUMN IF NOT EXISTS file_id INTEGER
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

  // Insert default config row if none exists
  const cfg = await pool.query('SELECT id FROM clinic_config LIMIT 1');
  if (cfg.rows.length === 0) {
    await pool.query('INSERT INTO clinic_config (configured, start_odip, next_odip) VALUES (false, 1, 1)');
  }

  // Insert default Sheet 1 if no sheets exist
  const fileRes = await pool.query('SELECT id FROM excel_files LIMIT 1');

let fileId;

if (fileRes.rows.length === 0) {
  const newFile = await pool.query(
    "INSERT INTO excel_files (name) VALUES ('Default File') RETURNING id"
  );
  fileId = newFile.rows[0].id;
} else {
  fileId = fileRes.rows[0].id;
}
ACTIVE_FILE_ID = fileId;

const sh = await pool.query('SELECT id FROM sheets LIMIT 1');

if (sh.rows.length === 0) {
  await pool.query(
    'INSERT INTO sheets (file_id, name, position) VALUES ($1, $2, $3)',
    [fileId, 'Sheet 1', 1]
  );
}

  console.log('  ✅  Database ready');
}

// ── All 17 Treatments ──────────────────────────────────────────────
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

// ── Load full state from DB ────────────────────────────────────────
async function loadState(fileId = ACTIVE_FILE_ID) {  const cfgRes = await pool.query('SELECT * FROM clinic_config LIMIT 1');
  const cfg    = cfgRes.rows[0];

const sheetsRes = await pool.query(
  'SELECT * FROM sheets WHERE file_id = $1 ORDER BY position ASC',
  [fileId]
);
  const sheets    = [];

  for (const sh of sheetsRes.rows) {
    const pRes = await pool.query(
      'SELECT * FROM patients WHERE sheet_id = $1 ORDER BY position ASC',
      [sh.id]
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
    configured: cfg.configured,
    startOdip:  cfg.start_odip,
    nextOdip:   cfg.next_odip,
    sheets
  };
}

// ── Renumber all ODIPs from scratch ───────────────────────────────
async function renumberOdips(startOdip, fileId = ACTIVE_FILE_ID) {
  const sheetsRes = await pool.query('SELECT id FROM sheets ORDER BY position ASC');
  let odip = startOdip;
  for (const sh of sheetsRes.rows) {
    const pRes = await pool.query(
      'SELECT id FROM patients WHERE sheet_id = $1 ORDER BY position ASC',
      [sh.id]
    );
    for (const p of pRes.rows) {
      await pool.query('UPDATE patients SET odip = $1 WHERE id = $2', [odip, p.id]);
      odip++;
    }
  }
  await pool.query('UPDATE clinic_config SET next_odip = $1', [odip]);
  return odip;
}

// ── Build Excel ────────────────────────────────────────────────────
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
    ['ODIP No.', 'Patient Name', 'Treatment Givent', 'Amount'].forEach((h, i) => {
      const c     = hRow.getCell(i + 1);
      c.value     = h;
      c.font      = { bold: true, size: 13, name: 'Calibri' };
      c.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    sheet.patients.forEach(p => {
      const row  = ws.addRow([p.odip, p.name, p.treatment, p.amount]);
      row.height = 18;
      row.eachCell((cell, col) => {
        cell.font      = { size: 11, name: 'Calibri' };
        cell.alignment = {
          vertical:   'middle',
          horizontal: col === 1 || col === 4 ? 'center' : 'left'
        };
      });
    });

    const totalRowNum   = sheet.patients.length + 2;
    const totalVal      = sheet.patients.reduce((s, p) => s + p.amount, 0);
    const tRow          = ws.getRow(totalRowNum);
    tRow.height         = 22;
    const lc            = tRow.getCell(3);
    lc.value            = 'Total';
    lc.font             = { bold: true, size: 14, name: 'Calibri' };
    lc.alignment        = { horizontal: 'center', vertical: 'middle' };
    const ac            = tRow.getCell(4);
    ac.value            = totalVal;
    ac.font             = { bold: true, size: 14, name: 'Calibri' };
    ac.alignment        = { horizontal: 'right', vertical: 'middle' };

    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }

  const buffer = await wb.xlsx.writeBuffer();
  return buffer;
}

// ── Middleware ─────────────────────────────────────────────────────
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ═══════════════════════════════════════════════════════════════════
//  API ROUTES
// ═══════════════════════════════════════════════════════════════════

// GET full state
app.get('/api/state', async (req, res) => {
  try {
    res.json(await loadState());
  } catch(e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// POST first-time ODIP setup (legacy - kept for compatibility)
app.post('/api/setup', async (req, res) => {
  const { startOdip } = req.body;
  if (!startOdip || startOdip < 1) return res.status(400).json({ error: 'Invalid ODIP' });
  try {
    await pool.query(
      'UPDATE clinic_config SET configured = true, start_odip = $1, next_odip = $1',
      [startOdip]
    );
    res.json({ ok: true });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// ═══════════════════════════════════════════════════════════════════
//  FEATURE 1: Auto New Session on Website Open
//  POST /api/session — creates new empty sheet, auto-increments position
// ═══════════════════════════════════════════════════════════════════
app.post('/api/session', async (req, res) => {
  try {

    const posRes = await pool.query(
      'SELECT COALESCE(MAX(position),0) as maxpos FROM sheets WHERE file_id=$1',
      [ACTIVE_FILE_ID]
    );

    const position = parseInt(posRes.rows[0].maxpos) + 1;

    const now = new Date();

    const name =
      now.getDate() + '-' +
      (now.getMonth()+1) + '-' +
      now.getFullYear();

    const shRes = await pool.query(
      'INSERT INTO sheets (file_id, name, position) VALUES ($1,$2,$3) RETURNING *',
      [ACTIVE_FILE_ID, name, position]
    );

    res.json({
      ok:true,
      sheetId:shRes.rows[0].id,
      sheetName:shRes.rows[0].name
    });

  } catch(e) {
    console.error(e);
    res.status(500).json({ error:e.message });
  }
});

// ═══════════════════════════════════════════════════════════════════
//  FEATURE 2: ODIP Control System
//  POST /api/odip/set — manually set starting ODIP
// ═══════════════════════════════════════════════════════════════════
app.post('/api/odip/set', async (req, res) => {
  const { odip } = req.body;
  if (!odip || odip < 1 || odip > 999999) {
    return res.status(400).json({ error: 'Invalid ODIP number' });
  }
  try {
    await pool.query(
      'UPDATE clinic_config SET next_odip = $1, configured = true',
      [odip]
    );
    res.json({ ok: true, nextOdip: odip });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// GET current ODIP info
app.get('/api/odip', async (req, res) => {
  try {
    const cfgRes = await pool.query('SELECT next_odip, start_odip FROM clinic_config LIMIT 1');
    const cfg = cfgRes.rows[0];
    res.json({ 
      nextOdip: cfg.next_odip, 
      startOdip: cfg.start_odip 
    });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// ═══════════════════════════════════════════════════════════════════
//  FEATURE 4: New Excel File Control
//  POST /api/excel/new — creates new Excel file (new DB state)
// ═══════════════════════════════════════════════════════════════════
app.post('/api/excel/new', async (req, res) => {

  try {

    const now = new Date();

    const fileName =
  req.body.name ||
  ('Excel File ' +
  now.getDate() + '-' +
  (now.getMonth()+1) + '-' +
  now.getFullYear());

    // create new excel file
    const fileRes = await pool.query(
      'INSERT INTO excel_files (name) VALUES ($1) RETURNING *',
      [fileName]
    );

    ACTIVE_FILE_ID = fileRes.rows[0].id;

    // create first sheet
    const shRes = await pool.query(
      'INSERT INTO sheets (file_id, name, position) VALUES ($1,$2,$3) RETURNING *',
      [ACTIVE_FILE_ID, 'Sheet 1', 1]
    );

    res.json({
      ok:true,
      fileId:ACTIVE_FILE_ID,
      sheetId:shRes.rows[0].id,
      sheetName:shRes.rows[0].name
    });

  } catch(e) {
    console.error(e);
    res.status(500).json({ error:e.message });
  }

});
// POST add patient to selected sheet
app.post('/api/patient', async (req, res) => {
  const { name, sheetId } = req.body;
  if (!name || !name.trim()) return res.status(400).json({ error: 'Name is required' });
  try {
    const shRes = await pool.query(
      'SELECT * FROM sheets WHERE id = $1',
      [sheetId]
    );
    const sheet = shRes.rows[0];
    if (!sheet) return res.status(404).json({ error: 'Sheet not found' });

    const cfgRes = await pool.query('SELECT next_odip FROM clinic_config LIMIT 1');
    const nextOdip = cfgRes.rows[0].next_odip;

    const countRes = await pool.query('SELECT COUNT(*) FROM patients WHERE sheet_id = $1', [sheet.id]);
    const position = parseInt(countRes.rows[0].count) + 1;

    const t = randTreatment();
    const pRes = await pool.query(
      'INSERT INTO patients (sheet_id, odip, name, treatment, amount, position) VALUES ($1,$2,$3,$4,$5,$6) RETURNING *',
      [sheet.id, nextOdip, name.trim(), t.name, t.amount, position]
    );
    const patient = pRes.rows[0];

    await pool.query('UPDATE clinic_config SET next_odip = next_odip + 1');
    const newNext = nextOdip + 1;

    res.json({
      patient: { id: patient.id, odip: patient.odip, name: patient.name, treatment: patient.treatment, amount: patient.amount },
      nextOdip: newNext
    });
  } catch(e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// DELETE last patient (undo)
app.delete('/api/patient/last', async (req, res) => {
  try {
    const { sheetId } = req.body;

const shRes = await pool.query(
  'SELECT * FROM sheets WHERE id = $1',
  [sheetId]
);

    const sheet = shRes.rows[0];

    const pRes = await pool.query(
      'SELECT * FROM patients WHERE sheet_id=$1 ORDER BY position DESC LIMIT 1',
      [sheet.id]
    );

    if(pRes.rows.length === 0){
      return res.status(400).json({ error:'No patients to undo' });
    }

    const patient = pRes.rows[0];
    await pool.query('DELETE FROM patients WHERE id=$1', [patient.id]);
    await pool.query('UPDATE clinic_config SET next_odip = next_odip - 1');
    const cfgRes = await pool.query('SELECT next_odip FROM clinic_config LIMIT 1');

    res.json({
      ok:true,
      removed:patient,
      nextOdip:cfgRes.rows[0].next_odip
    });
  } catch(e){
    console.error(e);
    res.status(500).json({ error:e.message });
  }
});

// PATCH rename sheet
app.patch('/api/sheet/:id/rename', async (req, res) => {
  const { name } = req.body;
  if(!name || !name.trim())
    return res.status(400).json({ error:'Name cannot be empty' });
  try{
    await pool.query('UPDATE sheets SET name=$1 WHERE id=$2', [name.trim(),req.params.id]);
    res.json({ ok:true, name:name.trim() });
  }catch(e){
    res.status(500).json({ error:e.message });
  }
});

// DELETE sheet
app.delete('/api/sheet/:id', async (req,res)=>{
  try{
    const id=parseInt(req.params.id);
    const countRes = await pool.query('SELECT COUNT(*) FROM sheets WHERE file_id=$1');
    if(parseInt(countRes.rows[0].count)<=1){
      return res.status(400).json({ error:'Cannot delete last sheet' });
    }
    await pool.query('DELETE FROM sheets WHERE id=$1', [id]);
    const cfg = await pool.query('SELECT start_odip FROM clinic_config LIMIT 1');
    await renumberOdips(cfg.rows[0].start_odip);
    res.json({ ok:true });
  }catch(e){
    console.log(e);
    res.status(500).json({ error:e.message });
  }
});

// DELETE specific patient by DB id
app.delete('/api/patient/:id', async (req, res) => {
  try {
    const id = parseInt(req.params.id);
    await pool.query('DELETE FROM patients WHERE id = $1', [id]);
    const cfgRes = await pool.query('SELECT start_odip FROM clinic_config LIMIT 1');
    await renumberOdips(cfgRes.rows[0].start_odip);
    res.json({ ok: true });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// POST create new sheet (manual)
app.post('/api/sheet', async (req, res) => {

  try {

    const countRes = await pool.query(
      'SELECT COUNT(*) FROM sheets WHERE file_id=$1',
      [ACTIVE_FILE_ID]
    );

    const num = parseInt(countRes.rows[0].count) + 1;

    const shRes = await pool.query(
      'INSERT INTO sheets (file_id, name, position) VALUES ($1,$2,$3) RETURNING *',
      [ACTIVE_FILE_ID, 'Sheet ' + num, num]
    );

    res.json({
      sheetName: shRes.rows[0].name,
      sheetId: shRes.rows[0].id
    });

  } catch(e) {

    console.error(e);

    res.status(500).json({
      error:e.message
    });

  }

});

// GET all excel files
app.get('/api/files', async (req, res) => {

  try {

    const filesRes = await pool.query(
      'SELECT * FROM excel_files ORDER BY created_at DESC'
    );

    res.json(filesRes.rows);

  } catch(e) {

    res.status(500).json({
      error:e.message
    });

  }

});

// GET all excel files
app.get('/api/files', async (req, res) => {

  try {

    const filesRes = await pool.query(
      'SELECT * FROM excel_files ORDER BY created_at DESC'
    );

    res.json(filesRes.rows);

  } catch(e) {

    res.status(500).json({
      error: e.message
    });

  }

});

// SWITCH excel file
app.post('/api/files/switch', async (req, res) => {

  try {

    const { fileId } = req.body;

    ACTIVE_FILE_ID = fileId;

    const state = await loadState(fileId);

    res.json({
      ok: true,
      state
    });

  } catch(e) {

    res.status(500).json({
      error: e.message
    });

  }

});

// GET download Excel
app.get('/api/download', async (req, res) => {
  try {
const state = await loadState(ACTIVE_FILE_ID);

const buffer = await buildExcel(state);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="patients.xlsx"');
    res.send(buffer);
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// SWITCH active excel file
app.post('/api/files/switch', async (req, res) => {

  try {

    const { fileId } = req.body;

    if(!fileId){
      return res.status(400).json({
        error:'File ID required'
      });
    }

    ACTIVE_FILE_ID = fileId;

    const state = await loadState(fileId);

    res.json({
      ok:true,
      state
    });

  } catch(e){

    res.status(500).json({
      error:e.message
    });

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

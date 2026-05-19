const express  = require('express');
const ExcelJS  = require('exceljs');
const path     = require('path');
const { Pool } = require('pg');

const app  = express();
const PORT = process.env.PORT || 3000;

// ── PostgreSQL connection ──────────────────────────────────────────
// Render gives you a DATABASE_URL environment variable automatically
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL
    ? { rejectUnauthorized: false }
    : false
});

// ── Create tables if they don't exist ─────────────────────────────
async function initDB() {
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
      id         SERIAL PRIMARY KEY,
      name       TEXT NOT NULL,
      position   INTEGER NOT NULL,
      created_at TIMESTAMP DEFAULT NOW()
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

  // Insert default config row if none exists
  const cfg = await pool.query('SELECT id FROM clinic_config LIMIT 1');
  if (cfg.rows.length === 0) {
    await pool.query('INSERT INTO clinic_config (configured, start_odip, next_odip) VALUES (false, 1, 1)');
  }

  // Insert default Sheet 1 if no sheets exist
  const sh = await pool.query('SELECT id FROM sheets LIMIT 1');
  if (sh.rows.length === 0) {
    await pool.query("INSERT INTO sheets (name, position) VALUES ('Sheet 1', 1)");
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
async function loadState() {
  const cfgRes = await pool.query('SELECT * FROM clinic_config LIMIT 1');
  const cfg    = cfgRes.rows[0];

  const sheetsRes = await pool.query('SELECT * FROM sheets ORDER BY position ASC');
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
async function renumberOdips(startOdip) {
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

    // Header — bold, size 13
    const hRow = ws.getRow(1);
    hRow.height = 22;
    ['ODIP No.', 'Patient Name', 'Treatment Givent', 'Amount'].forEach((h, i) => {
      const c     = hRow.getCell(i + 1);
      c.value     = h;
      c.font      = { bold: true, size: 13, name: 'Calibri' };
      c.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    // Data rows — plain, size 11
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

    // Total row — bold, size 14
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

  // Write to buffer and return — no file system needed
  const buffer = await wb.xlsx.writeBuffer();
  return buffer;
}

// ── Middleware ─────────────────────────────────────────────────────
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});
// ── API Routes ─────────────────────────────────────────────────────

// GET full state
app.get('/api/state', async (req, res) => {
  try {
    res.json(await loadState());
  } catch(e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// POST first-time ODIP setup
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

// POST add patient to last sheet
app.post('/api/patient', async (req, res) => {
  const { name } = req.body;
  if (!name || !name.trim()) return res.status(400).json({ error: 'Name is required' });
  try {
    // Get last sheet
    const shRes = await pool.query('SELECT * FROM sheets ORDER BY position DESC LIMIT 1');
    const sheet = shRes.rows[0];

    // Get current nextOdip
    const cfgRes = await pool.query('SELECT next_odip FROM clinic_config LIMIT 1');
    const nextOdip = cfgRes.rows[0].next_odip;

    // Count existing patients in this sheet for position
    const countRes = await pool.query('SELECT COUNT(*) FROM patients WHERE sheet_id = $1', [sheet.id]);
    const position = parseInt(countRes.rows[0].count) + 1;

    const t = randTreatment();
    const pRes = await pool.query(
      'INSERT INTO patients (sheet_id, odip, name, treatment, amount, position) VALUES ($1,$2,$3,$4,$5,$6) RETURNING *',
      [sheet.id, nextOdip, name.trim(), t.name, t.amount, position]
    );
    const patient = pRes.rows[0];

    // Increment nextOdip
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
    const shRes = await pool.query('SELECT * FROM sheets ORDER BY position DESC LIMIT 1');
    const sheet = shRes.rows[0];

// PATCH rename sheet
app.patch('/api/sheet/:id/rename', async (req, res) => {

const { name } = req.body;

if(!name || !name.trim())
return res.status(400).json({
error:'Name cannot be empty'
});

try{

await pool.query(
'UPDATE sheets SET name=$1 WHERE id=$2',
[name.trim(),req.params.id]
);

res.json({
ok:true,
name:name.trim()
});

}catch(e){

res.status(500).json({
error:e.message
});

}

});


// DELETE sheet
app.delete('/api/sheet/:id', async (req,res)=>{

try{

const id=parseInt(req.params.id);

const countRes=
await pool.query(
'SELECT COUNT(*) FROM sheets'
);

if(parseInt(
countRes.rows[0].count
)<=1){

return res.status(400).json({
error:'Cannot delete last sheet'
});

}

await pool.query(
'DELETE FROM sheets WHERE id=$1',
[id]
);

const cfg=
await pool.query(
'SELECT start_odip FROM clinic_config LIMIT 1'
);

await renumberOdips(
cfg.rows[0].start_odip
);

res.json({
ok:true
});

}catch(e){

console.log(e);

res.status(500).json({
error:e.message
});

}

});

// DELETE specific patient by DB id
app.delete('/api/patient/:id', async (req, res) => {
  try {
    const id = parseInt(req.params.id);
    await pool.query('DELETE FROM patients WHERE id = $1', [id]);

    // Re-number all ODIPs
    const cfgRes   = await pool.query('SELECT start_odip FROM clinic_config LIMIT 1');
    const startOdip = cfgRes.rows[0].start_odip;
    await renumberOdips(startOdip);

    res.json({ ok: true });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// POST create new sheet
app.post('/api/sheet', async (req, res) => {
  try {
    const countRes = await pool.query('SELECT COUNT(*) FROM sheets');
    const num      = parseInt(countRes.rows[0].count) + 1;
    const position = num;
    const shRes    = await pool.query(
      'INSERT INTO sheets (name, position) VALUES ($1, $2) RETURNING *',
      ['Sheet ' + num, position]
    );
    res.json({ sheetName: shRes.rows[0].name, sheetId: shRes.rows[0].id });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});


app.patch('/api/sheet/:id/rename', async (req, res) => {
  // DELETE sheet
app.delete('/api/sheet/:id', async (req, res) => {
  try {
    const id = parseInt(req.params.id);

    // Don't delete if only one sheet exists
    const countRes = await pool.query(
      'SELECT COUNT(*) FROM sheets'
    );

    if(parseInt(countRes.rows[0].count) <= 1){
      return res.status(400).json({
        error:'Cannot delete last sheet'
      });
    }

    await pool.query(
      'DELETE FROM sheets WHERE id=$1',
      [id]
    );

    // Re-number ODIPs
    const cfg = await pool.query(
      'SELECT start_odip FROM clinic_config LIMIT 1'
    );

    await renumberOdips(
      cfg.rows[0].start_odip
    );

    res.json({ok:true});

  } catch(e){
    res.status(500).json({
      error:e.message
    });
  }
});
  const { name } = req.body;
  if (!name || !name.trim()) return res.status(400).json({ error: 'Name cannot be empty' });
  try {
    await pool.query('UPDATE sheets SET name = $1 WHERE id = $2', [name.trim(), req.params.id]);
    res.json({ ok: true, name: name.trim() });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// GET download Excel
app.get('/api/download', async (req, res) => {
  try {
    const state  = await loadState();
    const buffer = await buildExcel(state);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="patients.xlsx"');
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

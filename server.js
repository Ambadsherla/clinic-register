const express = require('express');
const ExcelJS = require('exceljs');
const path    = require('path');
const fs      = require('fs');

const app  = express();
const PORT = process.env.PORT || 3000;

// ── Data folder ────────────────────────────────────────────────────
const dataDir  = path.join(__dirname, 'data');
if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });

const DATA_FILE  = path.join(dataDir, 'patients.json');
const EXCEL_FILE = path.join(dataDir, 'patients.xlsx');

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

// ── Load / Save ────────────────────────────────────────────────────
function loadData() {
  if (!fs.existsSync(DATA_FILE)) {
    const fresh = { configured: false, nextOdip: 1, startOdip: 1, sheets: [{ name: 'Sheet 1', patients: [] }] };
    fs.writeFileSync(DATA_FILE, JSON.stringify(fresh, null, 2));
    return fresh;
  }
  try {
    return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
  } catch(e) {
    const fresh = { configured: false, nextOdip: 1, startOdip: 1, sheets: [{ name: 'Sheet 1', patients: [] }] };
    fs.writeFileSync(DATA_FILE, JSON.stringify(fresh, null, 2));
    return fresh;
  }
}

function saveData(d) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(d, null, 2));
}

function randTreatment() {
  return TREATMENTS[Math.floor(Math.random() * TREATMENTS.length)];
}

// ── Build Excel ────────────────────────────────────────────────────
async function buildExcel(data) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Clinic Register';

  for (const sheet of data.sheets) {
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

    sheet.patients.forEach((p) => {
      const row = ws.addRow([p.odip, p.name, p.treatment, p.amount]);
      row.height = 18;
      row.eachCell((cell, col) => {
        cell.font      = { size: 11, name: 'Calibri' };
        cell.alignment = { vertical: 'middle', horizontal: col === 1 || col === 4 ? 'center' : 'left' };
      });
    });

    const totalRowNum   = sheet.patients.length + 2;
    const totalVal      = sheet.patients.reduce((s, p) => s + p.amount, 0);
    const tRow          = ws.getRow(totalRowNum);
    tRow.height         = 22;
    const labelCell     = tRow.getCell(3);
    labelCell.value     = 'Total';
    labelCell.font      = { bold: true, size: 14, name: 'Calibri' };
    labelCell.alignment = { horizontal: 'center', vertical: 'middle' };
    const amtCell       = tRow.getCell(4);
    amtCell.value       = totalVal;
    amtCell.font        = { bold: true, size: 14, name: 'Calibri' };
    amtCell.alignment   = { horizontal: 'right', vertical: 'middle' };

    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }

  await wb.xlsx.writeFile(EXCEL_FILE);
}

// ── Middleware ─────────────────────────────────────────────────────
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── API Routes ─────────────────────────────────────────────────────

app.get('/api/state', (req, res) => {
  res.json(loadData());
});

app.post('/api/setup', (req, res) => {
  const { startOdip } = req.body;
  if (!startOdip || startOdip < 1) return res.status(400).json({ error: 'Invalid ODIP' });
  const data      = loadData();
  data.configured = true;
  data.startOdip  = startOdip;
  data.nextOdip   = startOdip;
  saveData(data);
  res.json({ ok: true });
});

app.post('/api/patient', async (req, res) => {
  const { name } = req.body;
  if (!name || !name.trim()) return res.status(400).json({ error: 'Name is required' });
  const data    = loadData();
  const t       = randTreatment();
  const patient = { odip: data.nextOdip, name: name.trim(), treatment: t.name, amount: t.amount };
  data.sheets[data.sheets.length - 1].patients.push(patient);
  data.nextOdip++;
  saveData(data);
  await buildExcel(data);
  res.json({ patient, nextOdip: data.nextOdip });
});

app.delete('/api/patient/last', async (req, res) => {
  const data  = loadData();
  const sheet = data.sheets[data.sheets.length - 1];
  if (!sheet.patients.length) return res.status(400).json({ error: 'Nothing to undo' });
  const removed = sheet.patients.pop();
  data.nextOdip--;
  saveData(data);
  await buildExcel(data);
  res.json({ ok: true, removed, nextOdip: data.nextOdip });
});

app.delete('/api/patient/:si/:pi', async (req, res) => {
  const data = loadData();
  const si   = parseInt(req.params.si);
  const pi   = parseInt(req.params.pi);
  if (!data.sheets[si]) return res.status(404).json({ error: 'Sheet not found' });
  data.sheets[si].patients.splice(pi, 1);
  let odip = data.startOdip;
  for (const s of data.sheets) {
    for (const p of s.patients) { p.odip = odip++; }
  }
  data.nextOdip = odip;
  saveData(data);
  await buildExcel(data);
  res.json({ ok: true });
});

app.post('/api/sheet', async (req, res) => {
  const data = loadData();
  const num  = data.sheets.length + 1;
  data.sheets.push({ name: `Sheet ${num}`, patients: [] });
  saveData(data);
  await buildExcel(data);
  res.json({ sheetName: `Sheet ${num}`, sheetCount: data.sheets.length });
});

app.patch('/api/sheet/:si/rename', async (req, res) => {
  const data     = loadData();
  const si       = parseInt(req.params.si);
  const { name } = req.body;
  if (!data.sheets[si])      return res.status(404).json({ error: 'Sheet not found' });
  if (!name || !name.trim()) return res.status(400).json({ error: 'Name cannot be empty' });
  data.sheets[si].name = name.trim();
  saveData(data);
  await buildExcel(data);
  res.json({ ok: true, name: data.sheets[si].name });
});

app.get('/api/download', async (req, res) => {
  const data = loadData();
  await buildExcel(data);
  if (fs.existsSync(EXCEL_FILE)) {
    res.download(EXCEL_FILE, 'patients.xlsx');
  } else {
    res.status(400).json({ error: 'No data to export' });
  }
});

app.listen(PORT, () => {
  console.log('');
  console.log('  Clinic Register running!');
  console.log(`  Open: http://localhost:${PORT}`);
  console.log('');
});

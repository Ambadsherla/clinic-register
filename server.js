const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 3000;

// Make sure data folder exists
const dataDir = path.join(__dirname, 'data');
if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir);

const DATA_FILE   = path.join(dataDir, 'patients.json');
const EXCEL_FILE  = path.join(dataDir, 'patients.xlsx');

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

// ── Load / Save JSON data ──────────────────────────────────────────
function loadData() {
  if (!fs.existsSync(DATA_FILE)) {
    const fresh = { configured: false, nextOdip: 6759, startOdip: 6759, sheets: [{ name: 'Sheet 1', patients: [] }] };
    fs.writeFileSync(DATA_FILE, JSON.stringify(fresh, null, 2));
    return fresh;
  }
  return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
}

function saveData(d) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(d, null, 2));
}

function randTreatment() {
  return TREATMENTS[Math.floor(Math.random() * TREATMENTS.length)];
}

// ── Build Excel file — Black & White, Bold Header, Bold Total ─────
async function buildExcel(data) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Clinic Register';

  // All colours are black / white / grey only — no green, no colour
  const BLACK      = 'FF000000';
  const WHITE      = 'FFFFFFFF';
  const LIGHT_GREY = 'FFF2F2F2'; // very light grey for alternate rows
  const MID_GREY   = 'FFD9D9D9'; // border colour
  const DARK_GREY  = 'FF404040'; // total row background

  for (const sheet of data.sheets) {
    const ws = wb.addWorksheet(sheet.name);

    ws.getColumn(1).width = 12;
    ws.getColumn(2).width = 28;
    ws.getColumn(3).width = 50;
    ws.getColumn(4).width = 14;

    // ── Header row — bold, size 13, white text on black background ──
    const headerRow = ws.getRow(1);
    headerRow.height = 28;
    ['ODIP No.', 'Patient Name', 'Treatment Givent', 'Amount'].forEach((h, i) => {
      const cell = headerRow.getCell(i + 1);
      cell.value = h;
      cell.font      = { bold: true, size: 13, color: { argb: WHITE }, name: 'Calibri' };
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: BLACK } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border    = { bottom: { style: 'medium', color: { argb: MID_GREY } } };
    });

    // ── Data rows — normal size 11, no colour, thin grey border ────
    sheet.patients.forEach((p, i) => {
      const row = ws.addRow([p.odip, p.name, p.treatment, p.amount]);
      row.height = 20;
      // Alternate white and very light grey — no colour at all
      const bg = i % 2 === 0 ? WHITE : LIGHT_GREY;
      row.eachCell((cell, col) => {
        cell.font      = { size: 11, color: { argb: BLACK }, name: 'Calibri' };
        cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        cell.alignment = { vertical: 'middle', horizontal: col === 1 || col === 4 ? 'center' : 'left' };
        cell.border    = { bottom: { style: 'thin', color: { argb: MID_GREY } } };
      });
    });

    // ── TOTAL row — bold, size 13, white text on dark grey ─────────
    const total   = sheet.patients.reduce((s, p) => s + p.amount, 0);
    const dataEnd = sheet.patients.length + 1;
    ws.mergeCells(dataEnd + 1, 1, dataEnd + 1, 3);
    const totalRow = ws.getRow(dataEnd + 1);
    totalRow.height = 26;

    // "TOTAL" label (merged A–C)
    totalRow.getCell(1).value     = 'TOTAL';
    totalRow.getCell(1).font      = { bold: true, size: 13, color: { argb: WHITE }, name: 'Calibri' };
    totalRow.getCell(1).fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_GREY } };
    totalRow.getCell(1).alignment = { horizontal: 'right', vertical: 'middle' };
    totalRow.getCell(1).border    = { top: { style: 'medium', color: { argb: BLACK } } };

    // Amount value (column D)
    totalRow.getCell(4).value     = total;
    totalRow.getCell(4).font      = { bold: true, size: 13, color: { argb: WHITE }, name: 'Calibri' };
    totalRow.getCell(4).fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_GREY } };
    totalRow.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
    totalRow.getCell(4).border    = { top: { style: 'medium', color: { argb: BLACK } } };

    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }

  await wb.xlsx.writeFile(EXCEL_FILE);
}

// ── Middleware ─────────────────────────────────────────────────────
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── API Routes ─────────────────────────────────────────────────────

// Get full state
app.get('/api/state', (req, res) => {
  res.json(loadData());
});

// First-time ODIP setup
app.post('/api/setup', (req, res) => {
  const { startOdip } = req.body;
  if (!startOdip || startOdip < 1) return res.status(400).json({ error: 'Invalid ODIP' });
  const data = loadData();
  data.configured = true;
  data.startOdip  = startOdip;
  data.nextOdip   = startOdip;
  saveData(data);
  res.json({ ok: true });
});

// Add patient
app.post('/api/patient', async (req, res) => {
  const { name } = req.body;
  if (!name || !name.trim()) return res.status(400).json({ error: 'Name is required' });

  const data = loadData();
  const t = randTreatment();
  const patient = { odip: data.nextOdip, name: name.trim(), treatment: t.name, amount: t.amount };

  data.sheets[data.sheets.length - 1].patients.push(patient);
  data.nextOdip++;
  saveData(data);
  await buildExcel(data);

  res.json({ patient, nextOdip: data.nextOdip });
});

// Undo last patient
app.delete('/api/patient/last', async (req, res) => {
  const data = loadData();
  const sheet = data.sheets[data.sheets.length - 1];
  if (!sheet.patients.length) return res.status(400).json({ error: 'Nothing to undo' });

  const removed = sheet.patients.pop();
  data.nextOdip--;
  saveData(data);
  await buildExcel(data);
  res.json({ ok: true, removed, nextOdip: data.nextOdip });
});

// Delete specific patient
app.delete('/api/patient/:si/:pi', async (req, res) => {
  const data = loadData();
  const si = parseInt(req.params.si);
  const pi = parseInt(req.params.pi);
  if (!data.sheets[si]) return res.status(404).json({ error: 'Sheet not found' });

  data.sheets[si].patients.splice(pi, 1);

  // Re-number all ODIPs from start
  let odip = data.startOdip;
  for (const s of data.sheets) {
    for (const p of s.patients) { p.odip = odip++; }
  }
  data.nextOdip = odip;
  saveData(data);
  await buildExcel(data);
  res.json({ ok: true });
});

// Create new sheet
app.post('/api/sheet', async (req, res) => {
  const data = loadData();
  const num  = data.sheets.length + 1;
  data.sheets.push({ name: `Sheet ${num}`, patients: [] });
  saveData(data);
  await buildExcel(data);
  res.json({ sheetName: `Sheet ${num}`, sheetCount: data.sheets.length });
});

// Rename a sheet
app.patch('/api/sheet/:si/rename', async (req, res) => {
  const data = loadData();
  const si   = parseInt(req.params.si);
  const { name } = req.body;
  if (!data.sheets[si]) return res.status(404).json({ error: 'Sheet not found' });
  if (!name || !name.trim()) return res.status(400).json({ error: 'Name cannot be empty' });
  data.sheets[si].name = name.trim();
  saveData(data);
  await buildExcel(data);
  res.json({ ok: true, name: data.sheets[si].name });
});

// Download Excel
app.get('/api/download', async (req, res) => {
  const data = loadData();
  await buildExcel(data);
  res.download(EXCEL_FILE, 'patients.xlsx');
});

// Start server
app.listen(PORT, () => {
  console.log('');
  console.log('  ✅  Clinic Register is running!');
  console.log(`  👉  Open this in your browser: http://localhost:${PORT}`);
  console.log('');
  console.log('  Press Ctrl+C to stop the server.');
  console.log('');
});

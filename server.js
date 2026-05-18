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

// ── Build Excel — Plain, clean, no background colours ─────────────
async function buildExcel(data) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Clinic Register';

  for (const sheet of data.sheets) {
    const ws = wb.addWorksheet(sheet.name);

    // Column widths
    ws.getColumn(1).width = 12;  // ODIP No.
    ws.getColumn(2).width = 28;  // Patient Name
    ws.getColumn(3).width = 52;  // Treatment Givent
    ws.getColumn(4).width = 12;  // Amount

    // ── Row 1: Header — bold, size 13, NO background fill ──────────
    const hRow = ws.getRow(1);
    hRow.height = 22;
    ['ODIP No.', 'Patient Name', 'Treatment Givent', 'amount'].forEach((h, i) => {
      const c = hRow.getCell(i + 1);
      c.value = h;
      c.font      = { bold: true, size: 13, name: 'Calibri' };
      c.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    // ── Data rows — plain, size 11, no fill, no border ─────────────
    sheet.patients.forEach((p) => {
      const row = ws.addRow([p.odip, p.name, p.treatment, p.amount]);
      row.height = 18;
      row.eachCell((cell, col) => {
        cell.font      = { size: 11, name: 'Calibri' };
        cell.alignment = {
          vertical: 'middle',
          horizontal: col === 1 || col === 4 ? 'center' : 'left'
        };
      });
    });

    // ── TOTAL row — bold, size 14, plain white background ──────────
    // "Total" label goes in column C, amount in column D
    // (matches exactly what you showed in the screenshot)
    const totalRowNum = sheet.patients.length + 2; // row after last patient
    const totalVal    = sheet.patients.reduce((s, p) => s + p.amount, 0);

    const tRow = ws.getRow(totalRowNum);
    tRow.height = 22;

    // Column C: "Total" label — bold, size 14
    const labelCell = tRow.getCell(3);
    labelCell.value     = 'Total';
    labelCell.font      = { bold: true, size: 14, name: 'Calibri' };
    labelCell.alignment = { horizontal: 'center', vertical: 'middle' };

    // Column D: amount value — bold, size 14
    const amtCell = tRow.getCell(4);
    amtCell.value     = totalVal;
    amtCell.font      = { bold: true, size: 14, name: 'Calibri' };
    amtCell.alignment = { horizontal: 'right', vertical: 'middle' };

    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }

  await wb.xlsx.writeFile(EXCEL_FILE);
}

// ── Middleware ─────────────────────────────────────────────────────
app.use(express.json());

app.use(express.static(path.join(__dirname, "public")));

app.get("*", (req,res)=>{
   res.sendFile(path.join(__dirname,"public","index.html"));
});
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

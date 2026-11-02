const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 5000;

// Keep uploaded files in memory so nothing is permanently saved to disk.
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 15 * 1024 * 1024
  },
  fileFilter: (req, file, cb) => {
    const allowed = /\.(xlsx|xls|xlsm)$/i.test(file.originalname);
    if (!allowed) {
      return cb(new Error("Only .xlsx, .xls and .xlsm files are allowed."));
    }
    cb(null, true);
  }
});

app.use(express.static(path.join(__dirname, "public")));

function columnNumberToLetter(n) {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function deleteColumns(ws, startIndex, count) {
  const range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
  if (startIndex > range.e.c) return;

  const actualCount = Math.min(count, range.e.c - startIndex + 1);

  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = startIndex; c <= range.e.c - actualCount; c++) {
      const from = XLSX.utils.encode_cell({ r, c: c + actualCount });
      const to = XLSX.utils.encode_cell({ r, c });

      if (ws[from]) {
        ws[to] = ws[from];
      } else {
        delete ws[to];
      }
    }

    for (let c = range.e.c - actualCount + 1; c <= range.e.c; c++) {
      delete ws[XLSX.utils.encode_cell({ r, c })];
    }
  }

  range.e.c -= actualCount;
  ws["!ref"] = XLSX.utils.encode_range(range);
}

function deleteColumnsFrom(ws, startIndex) {
  const range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
  if (startIndex > range.e.c) return;

  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = startIndex; c <= range.e.c; c++) {
      delete ws[XLSX.utils.encode_cell({ r, c })];
    }
  }

  range.e.c = startIndex - 1;
  ws["!ref"] = XLSX.utils.encode_range(range);
}

function unmergeDE(ws) {
  if (!Array.isArray(ws["!merges"])) return;

  ws["!merges"] = ws["!merges"].filter((merge) => {
    // Remove merges that involve D or E.
    return !(merge.s.c <= 4 && merge.e.c >= 3);
  });
}

function removeColumnEAndResizeD(ws) {
  deleteColumns(ws, 4, 1);

  // Give column D a sensible width.
  ws["!cols"] = ws["!cols"] || [];
  ws["!cols"][3] = {
    ...(ws["!cols"][3] || {}),
    wch: 18
  };
}

function normalizeHeader(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

const requiredHeaders = [
  "Sr No.",
  "Student Id",
  "Student Name",
  "Attendance Date",
  "Batch Details",
  "Module Name",
  "Session Name"
];

function findHeaderRowAndColumns(ws) {
  const data = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: "",
    raw: false
  });

  for (let r = 0; r < Math.min(data.length, 30); r++) {
    const row = data[r].map(normalizeHeader);
    const found = {};

    for (const header of requiredHeaders) {
      const index = row.indexOf(normalizeHeader(header));
      if (index !== -1) found[header] = index;
    }

    if (Object.keys(found).length >= 5) {
      return { rowIndex: r, columns: found, data };
    }
  }

  return null;
}

function rebuildRequiredColumns(ws) {
  const result = findHeaderRowAndColumns(ws);
  if (!result) return false;

  const { rowIndex, columns, data } = result;

  // If all requested columns were found, rebuild the sheet with exactly those columns.
  if (Object.keys(columns).length !== requiredHeaders.length) {
    return false;
  }

  const output = [];
  output.push(requiredHeaders);

  for (let r = rowIndex + 1; r < data.length; r++) {
    const sourceRow = data[r];
    const newRow = requiredHeaders.map((header) => sourceRow[columns[header]] ?? "");

    // Skip completely empty rows.
    if (newRow.some((v) => String(v).trim() !== "")) {
      output.push(newRow);
    }
  }

  const newWs = XLSX.utils.aoa_to_sheet(output);
  newWs["!cols"] = [
    { wch: 9 },
    { wch: 16 },
    { wch: 25 },
    { wch: 18 },
    { wch: 22 },
    { wch: 22 },
    { wch: 25 }
  ];

  // Replace worksheet content.
  Object.keys(ws).forEach((key) => delete ws[key]);
  Object.assign(ws, newWs);

  return true;
}

function processWorkbook(buffer) {
  const workbook = XLSX.read(buffer, {
    type: "buffer",
    cellDates: true
  });

  if (!workbook.SheetNames.length) {
    throw new Error("The Excel file does not contain a worksheet.");
  }

  const sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];

  // Step 1: Enable editing / remove worksheet protection where present.
  delete ws["!protect"];

  // Step 2: Delete columns C-I (3rd through 9th columns).
  deleteColumns(ws, 2, 7);

  // Step 3: Delete columns H-K after Step 2.
  deleteColumns(ws, 7, 4);

  // Step 4: Delete K onwards.
  deleteColumnsFrom(ws, 10);

  // Step 5: Unmerge D and E.
  unmergeDE(ws);

  // Step 6: Delete E and resize D.
  removeColumnEAndResizeD(ws);

  // Finally, if the expected headers are present, rebuild exactly the required 7 columns.
  rebuildRequiredColumns(ws);

  return XLSX.write(workbook, {
    bookType: "xlsx",
    type: "buffer"
  });
}

app.post("/api/process", upload.single("excelFile"), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: "Please upload an Excel file." });
    }

    const output = processWorkbook(req.file.buffer);

    const originalName = path.parse(req.file.originalname).name;
    const safeName = originalName.replace(/[^a-zA-Z0-9_-]/g, "_");

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${safeName}_processed.xlsx"`
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    res.send(output);
  } catch (error) {
    console.error(error);
    res.status(500).json({
      message: error.message || "Excel processing failed."
    });
  }
});

// ---------------------------------------------------------------------------
// SAR  =  user's SAR file + processed attendance file
//   adds  "Portal"  = COUNTIF(attendance Student Id column, SAR Student Id)
//   adds  "Diff"    = Class Attendance - Portal
// ---------------------------------------------------------------------------
const ExcelJS = require("exceljs");

const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];

function currentMonthYear() {
  const parts = new Intl.DateTimeFormat("en-US", { timeZone: "Asia/Karachi", year: "numeric", month: "numeric" }).formatToParts(new Date());
  const year = Number(parts.find((p) => p.type === "year").value);
  const month = Number(parts.find((p) => p.type === "month").value);
  return { year, month, monthName: MONTHS[month - 1] };
}

function httpError(status, message, extra) {
  const e = new Error(message);
  e.status = status;
  e.extra = extra || {};
  return e;
}

function cellText(cell) {
  let v = cell ? cell.value : null;
  if (v && typeof v === "object") {
    if (v.result !== undefined) v = v.result;
    else if (v.richText) v = v.richText.map((t) => t.text).join("");
    else if (v.text !== undefined) v = v.text;
    else if (v instanceof Date) v = v.toISOString().slice(0, 10);
  }
  return v === null || v === undefined ? "" : String(v).trim();
}

function cellNumber(cell) {
  let v = cell ? cell.value : null;
  if (v && typeof v === "object" && v.result !== undefined) v = v.result;
  const n = Number(v);
  return v === null || v === "" || v === undefined || isNaN(n) ? 0 : n;
}

async function loadWorkbook(buffer, label) {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(buffer);
  } catch (e) {
    throw httpError(422, `The ${label} could not be read. Please upload a valid .xlsx file.`);
  }
  if (!wb.worksheets.length) throw httpError(422, `The ${label} has no worksheet.`);
  return wb;
}

// Find the header row of the SAR sheet and guess the Student Id / Class Attendance columns.
function inspectSarSheet(ws) {
  const maxScan = Math.min(ws.rowCount, 40);
  let best = null;
  for (let r = 1; r <= maxScan; r++) {
    const row = ws.getRow(r);
    const headers = [];
    row.eachCell({ includeEmpty: false }, (cell, c) => {
      const t = cellText(cell);
      if (t) headers.push({ col: c, name: t });
    });
    const idCol = headers.find((h) => /^(student\s*id|studentid|std\s*id|stud\.?\s*id|roll\s*(no|number)?)$/i.test(h.name.replace(/[._]/g, " ").trim()))
      || headers.find((h) => /student.*id|id.*student/i.test(h.name));
    if (idCol) { best = { headerRow: r, headers, idCol: idCol.col }; break; }
    if (!best && headers.length >= 3) best = { headerRow: r, headers, idCol: null };
  }
  if (!best) throw httpError(422, "Could not find a header row in the SAR file.");
  const att =
    best.headers.find((h) => /class.*attend|attend.*class/i.test(h.name)) ||
    best.headers.find((h) => /class.*(day|present)|(day|present).*class/i.test(h.name)) ||
    best.headers.find((h) => /attend|present/i.test(h.name) && !/portal|diff/i.test(h.name));
  best.attCol = att ? att.col : null;
  return best;
}

function colLetter(n) { return columnNumberToLetter(n); }

async function inspectSar(buffer) {
  const wb = await loadWorkbook(buffer, "SAR file");
  const ws = wb.worksheets[0];
  const info = inspectSarSheet(ws);
  return {
    sheet: ws.name,
    headerRow: info.headerRow,
    columns: info.headers.map((h) => ({ index: h.col, letter: colLetter(h.col), name: h.name })),
    studentIdCol: info.idCol,
    classAttendanceCol: info.attCol
  };
}

// Reads the processed attendance file and returns its rows (all 7 required columns must be present).
async function readProcessedAttendance(buffer) {
  const wb = await loadWorkbook(buffer, "attendance file");
  const ws = wb.worksheets[0];
  const headerRow = ws.getRow(1);
  const idx = {};
  const missing = [];
  const present = new Map();
  headerRow.eachCell({ includeEmpty: false }, (cell, c) => present.set(normalizeHeader(cellText(cell)), c));
  for (const h of requiredHeaders) {
    const c = present.get(normalizeHeader(h));
    if (!c) missing.push(h); else idx[h] = c;
  }
  if (missing.length) {
    throw httpError(
      422,
      "The attendance file is not a processed file (missing columns: " + missing.join(", ") +
      "). Please create it on the Cleaner page first, then upload it here.",
      { goCleaner: true }
    );
  }
  const rows = [];
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    if (r === 1) return;
    const out = requiredHeaders.map((h) => cellText(row.getCell(idx[h])));
    if (out.some((v) => v !== "")) rows.push(out);
  });
  if (!rows.length) throw httpError(422, "The attendance file has no data rows.");
  return rows;
}

async function buildSAR(sarBuffer, attBuffer, opts) {
  const attRows = await readProcessedAttendance(attBuffer);

  const wb = await loadWorkbook(sarBuffer, "SAR file");
  const ws = wb.worksheets[0];
  const info = inspectSarSheet(ws);
  const idCol = Number(opts.studentIdCol) || info.idCol;
  const attCol = Number(opts.classAttendanceCol) || info.attCol;
  if (!idCol) throw httpError(422, "Student Id column was not found in the SAR file.", { needColumns: true });
  if (!attCol) throw httpError(422, "Class Attendance column was not found in the SAR file.", { needColumns: true });
  if (idCol === attCol) throw httpError(422, "Student Id and Class Attendance must be different columns.", { needColumns: true });

  // 1) Copy the processed attendance into a new sheet inside the SAR workbook,
  //    so the COUNTIF formulas keep working after download.
  let sheetName = "Attendance";
  let n = 2;
  while (wb.getWorksheet(sheetName)) sheetName = "Attendance_" + n++;
  const wa = wb.addWorksheet(sheetName);
  wa.addRow(requiredHeaders);
  const counts = new Map();
  for (const r of attRows) {
    wa.addRow(r);
    const key = r[1].toLowerCase();
    counts.set(key, (counts.get(key) || 0) + 1);
  }
  wa.getRow(1).font = { bold: true };
  wa.views = [{ state: "frozen", ySplit: 1 }];
  [9, 16, 28, 16, 20, 40, 28].forEach((w, i) => (wa.getColumn(i + 1).width = w));
  const lastAttRow = attRows.length + 1;
  const attIdRange = `'${sheetName}'!$B$2:$B$${lastAttRow}`;

  // 2) Add Portal + Diff columns after the last header column.
  const headerRow = ws.getRow(info.headerRow);
  let lastCol = 0;
  headerRow.eachCell({ includeEmpty: false }, (cell, c) => { if (c > lastCol) lastCol = c; });
  const portalCol = lastCol + 1, diffCol = lastCol + 2;

  const hStyle = headerRow.getCell(lastCol);
  for (const [c, title] of [[portalCol, "Portal"], [diffCol, "Diff"]]) {
    const cell = headerRow.getCell(c);
    cell.value = title;
    cell.style = JSON.parse(JSON.stringify(hStyle.style || {}));
    if (!cell.font || !cell.font.bold) cell.font = { ...(cell.font || {}), bold: true };
    ws.getColumn(c).width = 12;
  }
  headerRow.commit && headerRow.commit();

  const idL = colLetter(idCol), attL = colLetter(attCol), portalL = colLetter(portalCol);
  let filled = 0, notFound = 0, mismatched = 0;

  for (let r = info.headerRow + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const sid = cellText(row.getCell(idCol));
    if (!sid) continue;
    const portal = counts.get(sid.toLowerCase()) || 0;
    const cls = cellNumber(row.getCell(attCol));

    const pc = row.getCell(portalCol);
    pc.value = { formula: `COUNTIF(${attIdRange},${idL}${r})`, result: portal };
    const dc = row.getCell(diffCol);
    dc.value = { formula: `${attL}${r}-${portalL}${r}`, result: cls - portal };

    const ref = row.getCell(attCol);
    for (const c of [pc, dc]) {
      if (ref.style) c.style = JSON.parse(JSON.stringify(ref.style));
      c.numFmt = "0";
    }
    filled++;
    if (!portal) notFound++;
    if (cls - portal !== 0) mismatched++;
    row.commit && row.commit();
  }
  if (!filled) throw httpError(422, "No student rows were found below the header in the SAR file.");

  const now = currentMonthYear();
  const fileName = `SAR_${now.monthName}_${now.year}.xlsx`;
  const out = await wb.xlsx.writeBuffer();
  return {
    buffer: Buffer.from(out),
    fileName,
    stats: { students: filled, attendanceRows: attRows.length, notInPortal: notFound, withDifference: mismatched }
  };
}

const sarUpload = upload.fields([{ name: "sarFile", maxCount: 1 }, { name: "attendanceFile", maxCount: 1 }]);

app.get("/sar", (req, res) => res.sendFile(path.join(__dirname, "public", "sar.html")));

// Step 1: look at the SAR file and tell the page which columns were detected.
app.post("/api/sar/inspect", upload.single("sarFile"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ message: "Please upload the SAR file." });
    res.json(await inspectSar(req.file.buffer));
  } catch (error) {
    res.status(error.status || 500).json({ message: error.message || "Could not read the SAR file.", ...(error.extra || {}) });
  }
});

// Step 2: build the final SAR.
app.post("/api/sar", sarUpload, async (req, res) => {
  try {
    const sar = req.files && req.files.sarFile && req.files.sarFile[0];
    const att = req.files && req.files.attendanceFile && req.files.attendanceFile[0];
    if (!sar || !att) {
      return res.status(400).json({ message: "Please upload both files: the SAR file and the processed attendance file." });
    }
    const { buffer, fileName, stats } = await buildSAR(sar.buffer, att.buffer, {
      studentIdCol: req.body.studentIdCol,
      classAttendanceCol: req.body.classAttendanceCol
    });
    res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("X-SAR-File-Name", fileName);
    res.setHeader("X-SAR-Stats", JSON.stringify(stats));
    res.setHeader("Access-Control-Expose-Headers", "X-SAR-File-Name, X-SAR-Stats, Content-Disposition");
    res.send(buffer);
  } catch (error) {
    console.error(error);
    res.status(error.status || 500).json({ message: error.message || "SAR generation failed.", ...(error.extra || {}) });
  }
});

app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError) {
    if (err.code === "LIMIT_FILE_SIZE") {
      return res.status(413).json({
        message: "File is too large. Maximum size is 15 MB."
      });
    }
  }

  res.status(400).json({
    message: err.message || "Invalid request."
  });
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});

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
  "Module Sequence",
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
    { wch: 18 },
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

  // Finally, if the expected headers are present, rebuild exactly the required 8 columns.
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

const express = require("express");
const ExcelJS = require("exceljs");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const crypto = require("crypto");

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_DIR = path.resolve(process.env.DATA_DIR || path.join(__dirname, "data"));
const UPLOAD_DIR = path.resolve(process.env.UPLOAD_DIR || path.join(__dirname, "uploads"));
const EXCEL_FILE = path.join(DATA_DIR, "Ward47_Public_Problems.xlsx");

if (process.env.NODE_ENV === "production" && !process.env.ADMIN_KEY?.trim()) {
  throw new Error("ADMIN_KEY must be configured in production.");
}

fs.mkdirSync(DATA_DIR, { recursive: true });
fs.mkdirSync(UPLOAD_DIR, { recursive: true });

// Enable CORS for frontend development (e.g. VS Code Live Server at 127.0.0.1:5500)
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Methods", "GET, POST, PATCH, DELETE, OPTIONS");
  res.header("Access-Control-Allow-Headers", "Content-Type, x-admin-key");
  if (req.method === "OPTIONS") return res.sendStatus(200);
  next();
});

// Static assets: serve current directory as well as public & uploads
app.use("/data", (_req, res) => res.sendStatus(404));
app.use(express.static(__dirname));
app.use(express.static(path.join(__dirname, "public")));
app.use("/uploads", express.static(UPLOAD_DIR));
app.use(express.json({ limit: "50kb" }));

// Route to serve admin portal directly
app.get("/admin", (_req, res) => {
  res.sendFile(path.join(__dirname, "admin.html"));
});

const storage = multer.diskStorage({
  destination: (_req, _file, cb) => cb(null, UPLOAD_DIR),
  filename: (_req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    cb(null, `${Date.now()}-${crypto.randomBytes(8).toString("hex")}${ext}`);
  }
});
const upload = multer({
  storage,
  limits: { fileSize: 5 * 1024 * 1024, files: 1 },
  fileFilter: (_req, file, cb) => {
    const allowed = ["image/jpeg", "image/png", "image/webp"];
    if (allowed.includes(file.mimetype)) cb(null, true);
    else cb(new Error("Only JPG, PNG, or WEBP images are allowed."));
  }
});

async function getWorkbook() {
  const workbook = new ExcelJS.Workbook();
  let worksheet;
  if (fs.existsSync(EXCEL_FILE)) {
    await workbook.xlsx.readFile(EXCEL_FILE);
    worksheet = workbook.getWorksheet("Complaints");
  }
  if (!worksheet) {
    worksheet = workbook.addWorksheet("Complaints");
    worksheet.columns = [
      { header: "Complaint ID", key: "complaintId", width: 18 },
      { header: "Submitted At", key: "submittedAt", width: 24 },
      { header: "Citizen Name", key: "name", width: 28 },
      { header: "Mobile Number", key: "mobile", width: 18 },
      { header: "Area", key: "area", width: 25 },
      { header: "Street", key: "street", width: 25 },
      { header: "Problem Type", key: "problemType", width: 30 },
      { header: "Description", key: "description", width: 60 },
      { header: "Photo Filename", key: "photoFilename", width: 40 },
      { header: "Status", key: "status", width: 18 }
    ];
    worksheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    worksheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF123B68" } };
  }
  return { workbook, worksheet };
}

function getNextComplaintId(worksheet) {
  let maxIdNum = 0;
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const val = String(row.getCell(1).value || "");
    const match = val.match(/W47-(\d+)/i);
    if (match) {
      const n = parseInt(match[1], 10);
      if (n > maxIdNum) maxIdNum = n;
    }
  });

  const counterFile = path.join(DATA_DIR, ".id_counter");
  let lastCounter = 0;
  if (fs.existsSync(counterFile)) {
    try {
      lastCounter = parseInt(fs.readFileSync(counterFile, "utf8").trim(), 10) || 0;
    } catch (_) {}
  }

  const nextNum = Math.max(maxIdNum, lastCounter) + 1;
  fs.writeFileSync(counterFile, String(nextNum), "utf8");
  return `W47-${String(nextNum).padStart(5, "0")}`;
}

app.post("/api/complaints", upload.single("photo"), async (req, res) => {
  try {
    const { name, mobile, area, street, problemType, description } = req.body;
    const required = [name, mobile, area, street, problemType, description];
    if (required.some(value => typeof value !== "string" || !value.trim())) {
      return res.status(400).json({ message: "Please fill in all required fields." });
    }
    if (!/^[6-9]\d{9}$/.test(mobile.trim())) {
      return res.status(400).json({ message: "Enter a valid 10-digit Indian mobile number." });
    }
    if (name.trim().length > 100 || area.trim().length > 120 ||
        street.trim().length > 120 || description.trim().length > 2000 ||
        problemType.trim().length > 100) {
      return res.status(400).json({ message: "One or more fields exceed the allowed length." });
    }

    const { workbook, worksheet } = await getWorkbook();
    const complaintId = getNextComplaintId(worksheet);
    const newRow = worksheet.addRow({
      complaintId,
      submittedAt: new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" }),
      name: safeExcelText(name),
      mobile: mobile.trim(),
      area: safeExcelText(area),
      street: safeExcelText(street),
      problemType: safeExcelText(problemType),
      description: safeExcelText(description),
      photoFilename: req.file ? req.file.filename : "",
      status: "Received"
    });
    newRow.alignment = { vertical: "middle" };
    await workbook.xlsx.writeFile(EXCEL_FILE);
    return res.status(201).json({ message: "Complaint stored.", complaintId });
  } catch (error) {
    if (req.file) fs.unlink(req.file.path, () => {});
    console.error("Complaint save error:", error);
    return res.status(500).json({ message: "Could not save complaint. Please try again." });
  }
});

// Mitigate Excel formula injection from user-submitted text.
function safeExcelText(value) {
  const text = value.trim();
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

app.get("/api/health", (_req, res) => res.json({ status: "ok" }));

// Admin endpoints are protected by ADMIN_KEY (from environment or persistent local secret)
function getAdminKey() {
  const keyFile = path.join(DATA_DIR, ".admin_key");
  if (fs.existsSync(keyFile)) {
    return fs.readFileSync(keyFile, "utf8").trim();
  }
  return process.env.ADMIN_KEY || null;
}

function requireAdmin(req, res, next) {
  const suppliedKey = req.get("x-admin-key") || "";
  if (!suppliedKey) {
    return res.status(401).json({ message: "Admin authorization key is missing." });
  }

  let configuredKey = getAdminKey();
  // If no admin key has ever been configured, save the first key as the permanent admin secret
  if (!configuredKey) {
    fs.writeFileSync(path.join(DATA_DIR, ".admin_key"), suppliedKey, "utf8");
    return next();
  }

  const expected = Buffer.from(configuredKey);
  const actual = Buffer.from(suppliedKey);
  if (expected.length !== actual.length || !crypto.timingSafeEqual(expected, actual)) {
    return res.status(401).json({ message: "Invalid admin key." });
  }
  next();
}

// Test / Verify Admin Key
app.get("/api/admin/verify", requireAdmin, (_req, res) => {
  res.json({ success: true, message: "Admin access granted." });
});

// Update / Change Admin Password on server
app.post("/api/admin/change-password", (req, res) => {
  const currentKey = req.get("x-admin-key") || "";
  const { newPassword } = req.body || {};

  if (!newPassword || typeof newPassword !== "string" || newPassword.trim().length < 4) {
    return res.status(400).json({ message: "New password must be at least 4 characters." });
  }

  const configuredKey = getAdminKey();
  if (configuredKey) {
    const expected = Buffer.from(configuredKey);
    const actual = Buffer.from(currentKey);
    if (expected.length !== actual.length || !crypto.timingSafeEqual(expected, actual)) {
      return res.status(401).json({ message: "Current password is incorrect." });
    }
  }

  const keyFile = path.join(DATA_DIR, ".admin_key");
  fs.writeFileSync(keyFile, newPassword.trim(), "utf8");
  res.json({ success: true, message: "Password updated successfully on server." });
});

// Fetch all complaints for Admin Dashboard
app.get("/api/admin/complaints", requireAdmin, async (_req, res) => {
  try {
    if (!fs.existsSync(EXCEL_FILE)) return res.json({ complaints: [] });
    const { worksheet } = await getWorkbook();
    const complaints = [];
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const value = (col) => {
        const cell = row.getCell(col).value;
        if (cell && typeof cell === "object" && cell.text) return cell.text;
        return cell == null ? "" : String(cell);
      };
      complaints.push({
        complaintId: value(1),
        submittedAt: value(2),
        name: value(3),
        mobile: value(4),
        area: value(5),
        street: value(6),
        problemType: value(7),
        description: value(8),
        photoFilename: value(9),
        status: value(10) || "Received"
      });
    });
    res.json({ complaints: complaints.reverse() });
  } catch (error) {
    console.error("Admin read error:", error);
    res.status(500).json({ message: "Could not load complaints." });
  }
});

// Update complaint status
app.patch("/api/admin/complaints/:complaintId", requireAdmin, async (req, res) => {
  try {
    const allowedStatuses = ["Received", "Pending", "In Progress", "Completed", "Rejected"];
    const { status } = req.body || {};
    if (!allowedStatuses.includes(status)) {
      return res.status(400).json({ message: "Invalid complaint status." });
    }
    if (!fs.existsSync(EXCEL_FILE)) return res.status(404).json({ message: "Complaint not found." });
    const { workbook, worksheet } = await getWorkbook();
    let targetRow = null;
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1 && String(row.getCell(1).value || "") === req.params.complaintId) targetRow = row;
    });
    if (!targetRow) return res.status(404).json({ message: "Complaint not found." });
    targetRow.getCell(10).value = status;
    await workbook.xlsx.writeFile(EXCEL_FILE);
    res.json({ message: "Status updated in Excel sheet.", complaintId: req.params.complaintId, status });
  } catch (error) {
    console.error("Admin status update error:", error);
    res.status(500).json({ message: "Could not update complaint status." });
  }
});

// Bulk delete all Completed complaints and update Excel sheet
app.delete("/api/admin/complaints/bulk/completed", requireAdmin, async (_req, res) => {
  try {
    if (!fs.existsSync(EXCEL_FILE)) {
      return res.status(404).json({ message: "No complaints found to delete." });
    }

    const { workbook, worksheet } = await getWorkbook();
    const rowsToDelete = [];

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        const status = String(row.getCell(10).value || "").trim().toLowerCase();
        if (status === "completed") {
          rowsToDelete.push({
            rowNumber,
            complaintId: String(row.getCell(1).value || ""),
            photoFilename: String(row.getCell(9).value || "")
          });
        }
      }
    });

    if (rowsToDelete.length === 0) {
      return res.json({ success: true, message: "No completed complaints found to delete.", count: 0 });
    }

    // Delete associated photo files from disk
    rowsToDelete.forEach(item => {
      if (item.photoFilename) {
        const photoPath = path.join(UPLOAD_DIR, item.photoFilename);
        if (fs.existsSync(photoPath)) {
          fs.unlink(photoPath, () => {});
        }
      }
    });

    // Delete rows from bottom to top so row numbers don't shift
    rowsToDelete.sort((a, b) => b.rowNumber - a.rowNumber);
    rowsToDelete.forEach(item => {
      worksheet.spliceRows(item.rowNumber, 1);
    });

    await workbook.xlsx.writeFile(EXCEL_FILE);
    console.log(`Deleted ${rowsToDelete.length} completed complaints from Excel sheet.`);

    return res.json({
      success: true,
      message: `${rowsToDelete.length} completed complaints deleted and Excel sheet updated.`,
      count: rowsToDelete.length
    });
  } catch (error) {
    console.error("Bulk delete completed complaints error:", error);
    return res.status(500).json({ message: "Could not delete completed complaints from Excel sheet." });
  }
});

// Delete a single complaint by ID and update Excel sheet
app.delete("/api/admin/complaints/:complaintId", requireAdmin, async (req, res) => {
  try {
    const { complaintId } = req.params;
    if (!fs.existsSync(EXCEL_FILE)) {
      return res.status(404).json({ message: "No complaints found." });
    }

    const { workbook, worksheet } = await getWorkbook();
    let targetRowNumber = -1;
    let photoFilename = "";

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1 && String(row.getCell(1).value || "").trim() === complaintId.trim()) {
        targetRowNumber = rowNumber;
        photoFilename = String(row.getCell(9).value || "");
      }
    });

    if (targetRowNumber === -1) {
      return res.status(404).json({ message: "Complaint not found." });
    }

    // Delete associated photo file if present
    if (photoFilename) {
      const photoPath = path.join(UPLOAD_DIR, photoFilename);
      if (fs.existsSync(photoPath)) {
        fs.unlink(photoPath, () => {});
      }
    }

    // Remove row from worksheet and rewrite Excel
    worksheet.spliceRows(targetRowNumber, 1);
    await workbook.xlsx.writeFile(EXCEL_FILE);

    console.log(`Complaint ${complaintId} deleted from Excel sheet.`);
    return res.json({ success: true, message: "Complaint deleted and Excel sheet updated.", complaintId });
  } catch (error) {
    console.error("Delete complaint error:", error);
    return res.status(500).json({ message: "Could not delete complaint from Excel sheet." });
  }
});

// Download Excel file directly from Admin Dashboard
app.get("/api/admin/export-excel", requireAdmin, async (_req, res) => {
  try {
    if (!fs.existsSync(EXCEL_FILE)) {
      return res.status(404).json({ message: "No complaints found to download." });
    }
    res.download(EXCEL_FILE, "Ward47_Public_Problems.xlsx");
  } catch (error) {
    console.error("Download error:", error);
    res.status(500).json({ message: "Could not download Excel file." });
  }
});

app.use((error, _req, res, _next) => {
  if (error instanceof multer.MulterError && error.code === "LIMIT_FILE_SIZE") {
    return res.status(400).json({ message: "Photo exceeds the 5 MB limit." });
  }
  return res.status(400).json({ message: error.message || "Invalid request." });
});

app.listen(PORT, () => {
  console.log(`Ward 47 Citizen Portal running at http://localhost:${PORT}`);
  console.log(`Admin Dashboard URL: http://localhost:${PORT}/admin`);
  console.log(`Excel complaints file: ${EXCEL_FILE}`);
});

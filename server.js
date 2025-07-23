// server.js
import express from "express";
import multer from "multer";
import path from "path";
import fs from "fs";
import XLSX from "xlsx";
import { fileURLToPath } from "url"; // Import fileURLToPath
import cors from "cors";

// Derive __dirname equivalent for ES Modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = 3000;

// Create an 'uploads' directory if it doesn't exist to store temporary files
const uploadsDir = path.join(__dirname, "uploads");
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir);
}

// Create an 'output' directory if it doesn't exist to store generated Excel files
const outputDir = path.join(__dirname, "output");
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir);
}

// Configure Multer for file uploads
// Files will be stored in the 'uploads' directory
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, uploadsDir);
  },
  filename: function (req, file, cb) {
    cb(
      null,
      file.fieldname + "-" + Date.now() + path.extname(file.originalname)
    );
  },
});

// Filter to allow only Excel files
const fileFilter = (req, file, cb) => {
  const allowedTypes = /xlsx|xls/; // Only Excel files now
  const extname = allowedTypes.test(
    path.extname(file.originalname).toLowerCase()
  );
  const mimetype = allowedTypes.test(file.mimetype);

  if (extname && mimetype) {
    return cb(null, true);
  } else {
    cb(new Error("Only Excel (.xlsx, .xls) files are allowed!"), false);
  }
};

const upload = multer({
  storage: storage,
  fileFilter: fileFilter,
  limits: { fileSize: 100 * 1024 * 1024 }, // 100 MB file size limit
});

// Serve static files from the 'public' directory
app.use(express.static(path.join(__dirname, "public")));
app.use("/output", express.static(outputDir)); // Serve generated files from /output URL
app.use(cors()); // Add this line before your routes

// --- Helper Functions ---

/**
 * Formats a JavaScript Date object into a YYYY-MM-DD string.
 * @param {Date} date The Date object to format.
 * @returns {string} The formatted date string.
 */
function formatDate(date) {
  if (!(date instanceof Date) || isNaN(date)) {
    return ""; // Return empty string for invalid dates
  }
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0"); // Months are 0-indexed
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

/**
 * Reads the content of an Excel file.
 * @param {string} filePath The path to the file.
 * @returns {Array<Object>} An array of objects, where each object is a row.
 */
function readExcelFile(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }
  const workbook = XLSX.readFile(filePath, { cellDates: true });
  const sheetName = workbook.SheetNames[0]; // Get the first sheet
  const worksheet = workbook.Sheets[sheetName];
  // Convert sheet to JSON array of objects (each object is a row) with header: 1 to get raw data and then map
  const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

  if (json.length === 0) {
    return []; // Empty Excel file
  }

  const headers = json[0];
  const rows = json.slice(1).map((row) => {
    const obj = {};
    headers.forEach((header, index) => {
      let value = row[index];
      if (value instanceof Date) {
        value = formatDate(value);
      }
      obj[header] = value;
    });
    return obj;
  });
  return rows;
}

/**
 * Generates and saves an Excel file from an array of data.
 * @param {Array<Object>} data The data to write to the Excel file.
 * @param {string} filename The name of the file to save.
 * @param {string} highlightColumn The name of the column to highlight (optional).
 */
function saveExcelFile(data, filename, highlightColumn = "") {
  let ws = XLSX.utils.json_to_sheet(data);

  if (highlightColumn && ws["!ref"]) {
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const headerRow = data[0] ? Object.keys(data[0]) : [];
    const colIndexToHighlight = headerRow.indexOf(highlightColumn);

    if (colIndexToHighlight !== -1) {
      const headerCellRef = XLSX.utils.encode_cell({
        r: 0,
        c: colIndexToHighlight,
      });
      if (ws[headerCellRef]) {
        ws[headerCellRef].c = ws[headerCellRef].c || [];
        ws[headerCellRef].c.push({ t: "Comparison Key", a: "Node.js Express" });
      }
    }
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Contents");
  XLSX.writeFile(wb, path.join(outputDir, filename)); // Save to output directory
  return path.join("/output", filename); // Return URL path for download
}

// --- Express Routes ---

// Route to handle file uploads and perform cross-check
app.post(
  "/cross-check",
  upload.fields([
    { name: "fileA", maxCount: 1 },
    { name: "fileB", maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      const fileA = req.files["fileA"] ? req.files["fileA"][0] : null;
      const fileB = req.files["fileB"] ? req.files["fileB"][0] : null;

      if (!fileA || !fileB) {
        return res.status(400).json({
          success: false,
          message: "Please upload both File A and File B.",
        });
      }

      // Ensure both files are Excel (Multer filter already handles this, but good to double check)
      const isFile1Excel =
        fileA.originalname.endsWith(".xlsx") ||
        fileA.originalname.endsWith(".xls");
      const isFile2Excel =
        fileB.originalname.endsWith(".xlsx") ||
        fileB.originalname.endsWith(".xls");

      if (!(isFile1Excel && isFile2Excel)) {
        // This case should ideally be caught by multer's fileFilter, but as a fallback
        return res.status(400).json({
          success: false,
          message: "Both files must be Excel files for comparison.",
        });
      }

      const data1 = readExcelFile(fileA.path);
      const data2 = readExcelFile(fileB.path);

      if (data1.length === 0) {
        // Clean up uploaded files
        fs.unlinkSync(fileA.path);
        fs.unlinkSync(fileB.path);
        return res.status(200).json({
          success: true,
          message: "File A is empty. Nothing to cross-check.",
          foundCount: 0,
          missingCount: 0,
          missingContents: [],
          matchedDownloadUrl: null,
          missingDownloadUrl: null,
        });
      }
      if (data2.length === 0) {
        // Clean up uploaded files
        fs.unlinkSync(fileA.path);
        fs.unlinkSync(fileB.path);
        return res.status(200).json({
          success: true,
          message: "File B is empty. No contents to compare against.",
          foundCount: 0,
          missingCount: 0,
          missingContents: [],
          matchedDownloadUrl: null,
          missingDownloadUrl: null,
        });
      }

      // Determine the comparison column: Use the first header from File A
      const comparisonColumn = Object.keys(data1[0])[0]; // Get the first header
      if (!comparisonColumn) {
        // Clean up uploaded files
        fs.unlinkSync(fileA.path);
        fs.unlinkSync(fileB.path);
        return res.status(400).json({
          success: false,
          message:
            "File A must have at least one column to perform comparison.",
        });
      }

      const values2Set = new Set(
        data2
          .map((row) => row[comparisonColumn]) // Use the first column for comparison
          .filter((v) => v !== undefined && v !== null && v !== "")
      );

      const foundInFile2 = [];
      const missingInFile2 = [];

      for (const row1 of data1) {
        const value1 = row1[comparisonColumn]; // Use the first column for comparison
        if (value1 !== undefined && value1 !== null && value1 !== "") {
          if (values2Set.has(value1)) {
            foundInFile2.push(row1); // Store the entire row if matched
          } else {
            missingInFile2.push(row1); // Store the entire row if missing
          }
        } else {
          // Consider rows with empty comparison column as missing for comparison
          missingInFile2.push(row1);
        }
      }

      const timestamp = new Date().toISOString().slice(0, 10);
      let matchedDownloadUrl = null;
      let missingDownloadUrl = null;

      if (foundInFile2.length > 0) {
        const matchedFilename = `matched_contents_${timestamp}.xlsx`;
        matchedDownloadUrl = saveExcelFile(
          foundInFile2,
          matchedFilename,
          comparisonColumn
        ); // Pass comparisonColumn for highlighting
      }

      if (missingInFile2.length > 0) {
        const missingFilename = `missing_contents_${timestamp}.xlsx`;
        missingDownloadUrl = saveExcelFile(
          missingInFile2,
          missingFilename,
          comparisonColumn
        ); // Pass comparisonColumn for highlighting
      }

      // Clean up uploaded files
      fs.unlinkSync(fileA.path);
      fs.unlinkSync(fileB.path);

      res.json({
        success: true,
        message: "Cross-check completed successfully!",
        foundCount: foundInFile2.length,
        missingCount: missingInFile2.length,
        totalFile1Rows: data1.length,
        missingContents: missingInFile2.slice(0, 10), // Send a sample for display
        matchedDownloadUrl: matchedDownloadUrl,
        missingDownloadUrl: missingDownloadUrl,
        file1Name: fileA.originalname,
        file2Name: fileB.originalname,
        comparisonColumn: comparisonColumn, // Inform the client which column was used
      });
    } catch (error) {
      console.error("Server error during cross-check:", error);
      // Clean up uploaded files in case of error
      if (req.files && req.files["fileA"] && req.files["fileA"][0])
        fs.unlinkSync(req.files["fileA"][0].path);
      if (req.files && req.files["fileB"] && req.files["fileB"][0])
        fs.unlinkSync(req.files["fileB"][0].path);
      res.status(500).json({
        success: false,
        message: `An error occurred: ${error.message}`,
      });
    }
  }
);

// Remove the /get-headers route as it's no longer needed.

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
  console.log(`Open your browser to http://localhost:${PORT}`);
});
//end

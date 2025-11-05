import { google } from "googleapis";
import { v4 as uuidv4 } from "uuid";
import path from "path";
import fs from "fs";
import XLSX from "xlsx";

const CREDENTIALS_PATH = "./google.json";
const SCOPES = ["https://www.googleapis.com/auth/spreadsheets"];
const SPREADSHEET_ID = "1ZwwBDyHc6bJunW9V_Uia3BzTrn5ahd8bkEjwc9xIUT0";

function generateUniqueId() {
  const now = Date.now().toString(36);
  const rand = Math.random().toString(36).substring(2, 4);
  return (now + rand).toUpperCase().slice(-6);
}

export const createUser = async (req, res) => {
  try {
    if (!req.files || !req.files.file) {
      return res.status(400).json({ message: "No file uploaded" });
    }

    const file = req.files.file;
    const uploadPath = path.join(process.cwd(), "uploads", file.name);

    if (!fs.existsSync(path.dirname(uploadPath))) {
      fs.mkdirSync(path.dirname(uploadPath), { recursive: true });
    }

    file.mv(uploadPath, async (err) => {
      if (err) {
        console.error("File upload error:", err);
        return res.status(500).json({
          message: "File upload failed",
          error: err.message,
        });
      }

      try {
        // 🔹 Read workbook
        const workbook = XLSX.readFile(uploadPath);
        const sheetNames = workbook.SheetNames;
        const firstSheet = workbook.Sheets[sheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });

        // 🔹 Split into header & values
        const headers = Object.keys(jsonData[0]);
        const values = jsonData.map((row) => Object.values(row));

        // 🔹 Combine for Excel export
        const finalExcelData = values

        // 🔹 Write to a new local Excel
        const newWorkbook = XLSX.utils.book_new();
        const newSheet = XLSX.utils.aoa_to_sheet(finalExcelData);
        XLSX.utils.book_append_sheet(newWorkbook, newSheet, "FormattedData");

        const newFilePath = path.join(
          process.cwd(),
          "uploads",
          "splitHeaderValue.xlsx"
        );
        XLSX.writeFile(newWorkbook, newFilePath);

        // 🔹 Create sheet in Google Sheets
        const response = await createEmpSheet(jsonData);

        return res.status(201).json({
          message: "Excel processed and uploaded successfully!",
          headers,
          sampleRow: values[0],
          googleStatus: response,
          newFilePath,
        });
      } catch (readErr) {
        console.error("Excel read error:", readErr);
        res.status(500).json({
          message: "Error reading Excel file",
          error: readErr.message,
        });
      }
    });
  } catch (e) {
    console.error("Error creating user:", e);
    res.status(500).json({ message: "Error creating user", error: e.message });
  }
};

async function createEmpSheet(jsonData) {
  try {
    const auth = new google.auth.GoogleAuth({
      keyFile: CREDENTIALS_PATH,
      scopes: SCOPES,
    });

    const client = await auth.getClient();
    const sheets = google.sheets({ version: "v4", auth: client });

    const spreadsheet = await sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
    });

    const masterSheet = spreadsheet.data.sheets.find(
      (s) => s.properties.title === "master"
    );
    if (!masterSheet) throw new Error("❌ Master sheet not found!");

    const copyResponse = await sheets.spreadsheets.sheets.copyTo({
      spreadsheetId: SPREADSHEET_ID,
      sheetId: masterSheet.properties.sheetId,
      requestBody: { destinationSpreadsheetId: SPREADSHEET_ID },
    });

    const newSheetId = copyResponse.data.sheetId;
    const newSheetName = `EmpDetails-${generateUniqueId()}`;

    // Rename new sheet
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            updateSheetProperties: {
              properties: { sheetId: newSheetId, title: newSheetName },
              fields: "title",
            },
          },
        ],
      },
    });

    // Upload your Excel data
    const headers = Object.keys(jsonData[0]);
    const values = jsonData.map((obj) => headers.map((h) => obj[h]));
    const empDetails = [headers, ...values];

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${newSheetName}!A1:${String.fromCharCode(
        65 + headers.length - 1
      )}${empDetails.length}`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: empDetails },
    });

    // Copy down formulas
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            copyPaste: {
              source: {
                sheetId: newSheetId,
                startRowIndex: 1,
                endRowIndex: 2,
                startColumnIndex: 3,
                endColumnIndex: 9,
              },
              destination: {
                sheetId: newSheetId,
                startRowIndex: 2,
                endRowIndex: empDetails.length,
                startColumnIndex: 3,
                endColumnIndex: 9,
              },
              pasteType: "PASTE_FORMULA",
            },
          },
        ],
      },
    });

    return "Formulas copied down successfully!";
  } catch (error) {
    console.error("❌ Error:", error.message);
    return "Error while creating Google Sheet!";
  }
}

export const getUsers = async (req, res) => {
  try {
    const data = await readSheet();
    res.status(200).json({ message: "Data fetched ✅", data });
  } catch (error) {
    res.status(500).json({ error: "Something went wrong!" });
  }
};

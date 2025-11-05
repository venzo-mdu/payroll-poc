import { google } from "googleapis";
import { v4 as uuidv4 } from "uuid";
import path from "path";
import fs from "fs";
import XLSX from "xlsx";

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
        return res
          .status(500)
          .json({ message: "File upload failed", error: err.message });
      }

      try {
        const workbook = XLSX.readFile(uploadPath);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        const jsonData = XLSX.utils.sheet_to_json(sheet);
        const createNewExcel = await createEmpSheet(jsonData);
        // Send response
        if (createNewExcel === "Formulas copied down successfully!") {
          return res.status(201).json({
            message: "File uploaded and read successfully",
          });
        }
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

const CREDENTIALS_PATH = "./google.json";
const SCOPES = ["https://www.googleapis.com/auth/spreadsheets"];
// const SPREADSHEET_ID = "1L-ZQKG1Y3541s1xD9-Y_S_25Ta191oSbrw8uJHnq2oQ";
const SPREADSHEET_ID = "1ZwwBDyHc6bJunW9V_Uia3BzTrn5ahd8bkEjwc9xIUT0";

function generateUniqueId() {
  const now = Date.now().toString(36); // timestamp part
  const rand = Math.random().toString(36).substring(2, 4); // random part
  return (now + rand).substring(-6).toUpperCase().slice(-6);
}

const uuid = generateUniqueId();
const newExcelName = `EmpDetails-${uuid}`;

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
    const newSheetName = newExcelName;

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
    const headers = Object.keys(jsonData[0]);
    const values = jsonData.map((obj) => headers.map((h) => obj[h]));
    const empDetails = [headers, ...values];

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${newSheetName}!A1:C${empDetails.length}`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: empDetails },
    });

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

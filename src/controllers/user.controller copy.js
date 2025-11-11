import { google } from "googleapis";
import dotenv from "dotenv";
dotenv.config();
import path from "path";
import fs from "fs";
import XLSX from "xlsx";
import { parseSalarySheet } from "./salaryExcel.js";
import {
  findVal,
  generateUniqueId,
  requiredHeaders,
} from "./generateUniqueId.js";
import fetch from "node-fetch";

const CREDENTIALS_PATH = process.env.CREDENTIALS_PATH;
const SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/drive.readonly",
];
const SPREADSHEET_ID = "1ZwwBDyHc6bJunW9V_Uia3BzTrn5ahd8bkEjwc9xIUT0";

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

        const salarySheet = workbook.SheetNames[1];

        const sheetSalary = workbook.Sheets[salarySheet];

        const jsonDataSalary = XLSX.utils.sheet_to_json(sheetSalary);

        const createNewExcel = await createEmpSheet(jsonData, jsonDataSalary);

        if (createNewExcel) {
          downloadExcelDirectly(SPREADSHEET_ID, createNewExcel);
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

async function createEmpSheet(jsonData, jsonDataSalary) {
  try {
    const uuid = generateUniqueId();
    const newExcelName = `EmpDetails-${uuid}`;

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

    const headerValues = Object.values(jsonData[0] || {});

    const headers = Object.keys(jsonData[0] || {});

    const cleanedHeaders = [...requiredHeaders];

    const cleanedHeaderKeys = requiredHeaders.map((reqHeader) => {
      let index = headerValues.indexOf(reqHeader);
      if (reqHeader === "Category" && index === -1) {
        index = headerValues.indexOf("Designation");
      }
      return index !== -1 ? headers[index] : null;
    });

    let totalIndex = findVal(jsonData, "Total");
    let conveyanceDaysIndex = findVal(jsonData, "Total") - 1;

    const values = jsonData.slice(1).map((rowObj) =>
      requiredHeaders.map((reqHeader, idx) => {
        const key = cleanedHeaderKeys[idx];

        const parsedData = parseSalarySheet(jsonDataSalary);

        let colName =
          rowObj.__EMPTY_3 === "Casual Labour"
            ? "Casual Labour (Unskilled)"
            : rowObj.__EMPTY_3 === "Jr. Supervisor"
            ? "Jr. Supervisor (Semi Skilled)"
            : rowObj.__EMPTY_3 === "Jr. Supervisor"
            ? "Jr. Supervisor (Semi Skilled)"
            : rowObj.__EMPTY_3 === "Tr. Supervisor"
            ? "Trainee Supervisor (Semi Skilled)"
            : rowObj.__EMPTY_3 === "Sr. MHE"
            ? "Jr. MHE Operator (Semi Skilled)"
            : rowObj.__EMPTY_3;

        if (reqHeader === "Fixed Basic") {
          return parsedData[colName] ? parsedData[colName]["BASIC WAGES"] : 0;
        }

        if (reqHeader === "Fixed VDA") {
          return parsedData[colName] ? parsedData[colName]["VDA"] : 0;
        }
        if (reqHeader === "HRA") {
          return parsedData[colName] ? parsedData[colName]["HRA"] : 0;
        }

        if (reqHeader === "Man Days") {
          const totalKey = `__EMPTY_${conveyanceDaysIndex}`;
          return (rowObj[totalKey] =
            rowObj[`__EMPTY_${conveyanceDaysIndex}`] || 0);
        }
        if (reqHeader === "Allowance Days") {
          const conveyanceDaysIndexKey = `__EMPTY_${totalIndex}`;
          return (rowObj[conveyanceDaysIndexKey] =
            rowObj[`__EMPTY_${totalIndex}`] || 0);
        }
        return key && rowObj[key] !== undefined && rowObj[key];
      })
    );

    const empDetails = [cleanedHeaders, ...values];

    console.log(empDetails);

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${newSheetName}!A1:AQ${empDetails.length}`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: empDetails },
    });

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            autoResizeDimensions: {
              dimensions: {
                sheetId: newSheetId,
                dimension: "COLUMNS",
                startIndex: 0,
                endIndex: empDetails.length,
              },
            },
          },
        ],
      },
    });

    return newSheetId;
  } catch (error) {
    console.error("❌ Error:", error.message);
  }
}

async function downloadExcelDirectly(SPREADSHEET_ID, GID) {
  try {
    const EXPORT_URL = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=xlsx&gid=${GID}`;

    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    const FILE_NAME = `EmployeeData_${GID}_${timestamp}.xlsx`;

    const response = await fetch(EXPORT_URL);
    if (!response.ok)
      throw new Error(`HTTP ${response.status} - ${response.statusText}`);

    const arrayBuffer = await response.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    const homeDir = process.env.USERPROFILE || process.env.HOME;
    const downloadsFolder = path.join(homeDir, "Downloads");

    if (!fs.existsSync(downloadsFolder)) {
      fs.mkdirSync(downloadsFolder, { recursive: true });
    }

    const destPath = path.join(downloadsFolder, FILE_NAME);
    fs.writeFileSync(destPath, buffer);
  } catch (err) {
    console.error("❌ Error downloading Excel:", err.message);
  }
}

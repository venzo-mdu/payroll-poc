import { google } from "googleapis";
import dotenv from "dotenv";
dotenv.config();
import path from "path";
import fs from "fs";
import XLSX from "xlsx";
import { parseSalarySheet } from "./salaryExcel.js";
import {
  calName,
  findVal,
  generateUniqueId,
  requiredHeaders,
} from "./generateUniqueId.js";

const CREDENTIALS_PATH = process.env.CREDENTIALS_PATH;
const SCOPES = [process.env.SCOPES];
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const uuid = generateUniqueId();
const newExcelName = `EmpDetails-${uuid}`;

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

async function createEmpSheet(jsonData, jsonDataSalary) {
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

    return "Formulas copied down successfully!";
  } catch (error) {
    console.error("❌ Error:", error.message);
  }
}

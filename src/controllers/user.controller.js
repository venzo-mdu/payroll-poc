import { google } from "googleapis";
import dotenv from "dotenv";
import path from "path";
import fs from "fs";
import XLSX from "xlsx";

dotenv.config();

const CREDENTIALS_PATH = process.env.CREDENTIALS_PATH;
const SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/drive",
];
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const addAttendanceSheetLength = async (workbook) => {
  const attendanceSheetName = workbook.SheetNames.find((name) =>
    name.toLowerCase().includes("attendance")
  );
  if (!attendanceSheetName) {
    console.log("❌ No sheet found containing 'Attendance'");
    return 0;
  }
  const attendanceSheet = workbook.Sheets[attendanceSheetName];
  const data = XLSX.utils.sheet_to_json(attendanceSheet, {
    header: 1,
    defval: "",
  });
  const dataWithoutHeader = data.slice(1);

  const attendanceUserLength = dataWithoutHeader.length;

  return attendanceUserLength;
};

export const createUser = async (req, res) => {
  try {
    if (!req.files || !req.files.file) {
      return res.status(400).json({
        success: false,
        message: "No file uploaded",
      });
    }

    const excelFile = req.files.file;
    const uploadPath = path.join("./uploads", excelFile.name);

    if (!fs.existsSync("./uploads")) fs.mkdirSync("./uploads");

    await excelFile.mv(uploadPath);

    // Read workbook
    const workbook = XLSX.readFile(uploadPath);
    const sheetNames = workbook.SheetNames;

    const auth = new google.auth.GoogleAuth({
      keyFile: CREDENTIALS_PATH,
      scopes: SCOPES,
    });
    const client = await auth.getClient();
    const sheets = google.sheets({ version: "v4", auth: client });

    const spreadsheet = await sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
    });

    const existingSheets = spreadsheet.data.sheets.map(
      (s) => s.properties.title
    );

    for (const sheetName of sheetNames) {
      if ("Rev Salary " === sheetName) {
        const searchKey = "Attendance";

        const result = Object.values(sheetNames).find((name) =>
          name.toLowerCase().includes(searchKey.toLowerCase())
        );

        console.log(result);
        let attendanceLength = await addAttendanceSheetLength(workbook);
        let s = workbook.Sheets[sheetName];
        const sheetData = XLSX.utils.sheet_to_json(s, {
          header: 1,
          defval: "",
        })[0];
        await uploadSalarySheet(
          workbook.Sheets[sheetName],
          sheets,
          existingSheets,
          SPREADSHEET_ID,
          sheetData,
          attendanceLength
        );
      } else {
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        if (!jsonData.length) continue;
        if (!existingSheets.includes(sheetName)) {
          await sheets.spreadsheets.batchUpdate({
            spreadsheetId: SPREADSHEET_ID,
            requestBody: {
              requests: [
                {
                  addSheet: {
                    properties: {
                      title: sheetName,
                    },
                  },
                },
              ],
            },
          });
          console.log(`✅ Created new sub-sheet: ${sheetName}`);
        } else {
          console.log(`ℹ️ Updating existing sub-sheet: ${sheetName}`);
        }
        await sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `${sheetName}!A1`,
          valueInputOption: "USER_ENTERED",
          requestBody: {
            values: jsonData,
          },
        });
      }
    }

    fs.unlinkSync(uploadPath);

    res.json({
      success: true,
      message: `✅ Uploaded ${sheetNames.length} sub-sheets to Google Sheets.`,
      uploadedSheets: sheetNames,
    });
  } catch (error) {
    console.error("❌ Error uploading Excel:", error.message);
    res.status(500).json({ success: false, error: error.message });
  }
};

const uploadSalarySheet = async (
  sheet,
  sheets,
  existingSheets,
  SPREADSHEET_ID,
  sheetData,
  attendanceLength
) => {
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const allDataRows = [];

  const FORMULA_SHIFT = 0;
  const EXCEL_HEADER_ROW = 0;

  for (let row = range.s.r; row <= range.e.r; row++) {
    const rowData = [];
    let rowHasFormula = false;

    const isHeader = row === -1;
    const isMetadata = row < -1;

    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: row + 1, c: col });
      const cell = sheet[cellAddress];

      if (!cell) {
        rowData.push("");
        continue;
      }

      if (cell.f) {
        let formula = cell.f;
        rowHasFormula = true;

        formula = formula.replace(/([A-Z]+)(\d+)/g, (match, column, rowNum) => {
          const newRowNum = parseInt(rowNum, 10) - FORMULA_SHIFT;
          return column + Math.max(1, newRowNum);
        });

        rowData.push(`=${formula}`);
      } else if (isHeader || isMetadata) {
        rowData.push(cell.v !== undefined ? cell.v : "");
      } else {
        rowData.push("");
      }
    }
    if (row < EXCEL_HEADER_ROW + 1 || rowHasFormula) {
      allDataRows.push(rowData);
    }
  }
  const dataForUpload = allDataRows.slice(FORMULA_SHIFT);

  const sheetName = "Rev Salary";

  if (!existingSheets.includes(sheetName)) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{ addSheet: { properties: { title: sheetName } } }],
      },
    });
    console.log(`✅ Created new sub-sheet: ${sheetName}`);
  } else {
    console.log(`ℹ️ Updating existing sub-sheet: ${sheetName}`);
  }

  // let result = [];

  // for (let i = 1; i < attendanceLength; i++) {
  //   let updatedRow = [...dataForUpload[0]];

  //   console.log(updatedRow);
  //   updatedRow[0] = 1;

  //   result.push(updatedRow);
  // }

  let result = [];

  result.push(dataForUpload[0]);

  for (let i = 1; i <= attendanceLength; i++) {
    let updatedRow = [...dataForUpload[0]];
    updatedRow = updatedRow.map((cell) => {
      if (typeof cell === "string" && cell.startsWith("=")) {
        return cell.replace(/([A-Z]+\$?)(\d+)/g, (match, col, rowNum) => {
          return `${col}${i + 1}`;
        });
      }
      return cell;
    });
    updatedRow[0] = i;

    result.push(updatedRow);
  }

  // const combinedData = [sheetData, ...dataForUpload.slice(10)];

  let d = dataForUpload[0];

  let [head, ...rest] = result

  const combinedData = [sheetData, ...rest];

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A1`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: combinedData },
  });
};

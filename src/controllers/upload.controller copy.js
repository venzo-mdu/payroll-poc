import path from "path";
import fs from "fs";
import bucket from "../Firebase/firebaseConfig.js";
import axios from "axios";
import XLSX from "xlsx";

export const ExcelToFirebase = async (filePath) => {
  try {
    const companyName = "TEN";
    if (!companyName) throw new Error("Company name required");

    const fileName = path.basename(filePath);
    const now = new Date();
    const year = now.getFullYear();
    const month = now.toLocaleString("default", { month: "long" });
    const destination = `${companyName}/${year}/${month}/${fileName}`;

    await bucket.upload(filePath, {
      destination,
      metadata: {
        contentType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
    });

    const fileUrl = `${process.env.FIREBASE_STORAGE_URL}/${bucket.name}/${destination}`;
    return fileUrl;
  } catch (error) {
    console.error("❌ Firebase upload error:", error);
    throw error;
  }
};

export const fetchFilesFromFirebase = async () => {
  try {
    const companyName = "TEN";
    const now = new Date();
    const year = now.getFullYear();
    const month = now.toLocaleString("default", { month: "long" });

    const prefix = `${companyName}/${year}/${month}/`;

    console.log("🔍 Searching in Firebase Storage path:", prefix);

    const [files] = await bucket.getFiles({ prefix });
    if (files.length === 0) {
      console.log(`⚠️ No files found in ${prefix}`);
      return [];
    }

    console.log(`✅ Found ${files.length} files:`);

    for (const f of files) console.log(" -", f.name);

    // Generate signed URLs (valid for 1 hour)
    const fileUrls = await Promise.all(
      files.map(async (file) => {
        const [signedUrl] = await file.getSignedUrl({
          action: "read",
          expires: Date.now() + 60 * 60 * 1000, // valid for 1 hour
        });
        return signedUrl;
      })
    );

    console.log("✅ Signed URLs generated:");
    fileUrls.forEach((url, i) => console.log(`${i + 1}. ${url}`));

    const firstFileUrl = fileUrls[0];
    console.log("⬇️ Downloading first file:", firstFileUrl);

    let result = await downloadExcelFromUrl(firstFileUrl);
    return result;
  } catch (error) {
    console.error("❌ Error fetching files from Firebase:", error);
    throw error;
  }
};

export const downloadExcelFromUrl = async (
  url,
  saveAs = "./downloads/file.xlsx"
) => {
  try {
    const response = await axios.get(url, {
      responseType: "arraybuffer",
      validateStatus: (status) => status < 500,
    });

    if (response.status !== 200) {
      throw new Error(`Failed to download file. Status: ${response.status}`);
    }

    fs.mkdirSync(path.dirname(saveAs), { recursive: true });
    fs.writeFileSync(saveAs, response.data);

    const workbook = XLSX.read(response.data, { type: "buffer" });
    console.log("✅ Excel file downloaded and read successfully!");
    return workbook;
  } catch (error) {
    console.error("❌ Error downloading Excel file:", error.message);
    throw error;
  }
};

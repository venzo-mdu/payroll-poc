import path from "path";
import fs from "fs";
import bucket from "../Firebase/firebaseConfig.js";

export const uploadExcelToFirebase = async (req, res) => {
  try {
     if (!req.files || !req.files.file) {
      return res.status(400).json({ success: false, message: "No file uploaded" });
    }

    const companyName = 'TEN'
    if (!companyName) {
      return res.status(400).json({ success: false, message: "Company name required" });
    }

    const file = req.files.file;
    const uploadPath = path.join("./uploads", file.name);

    if (!fs.existsSync("./uploads")) fs.mkdirSync("./uploads");
    await file.mv(uploadPath);

    const now = new Date();
    const year = now.getFullYear();
    const month = now.toLocaleString("default", { month: "long" });
    const destination = `${companyName}/${year}/${month}/${file.name}`;

    await bucket.upload(uploadPath, {
      destination,
      metadata: { contentType: file.mimetype },
    });

    fs.unlinkSync(uploadPath);

    const fileUrl = `${process.env.FIREBASE_STORAGE_URL}/${bucket.name}/${destination}`;

    return res.status(200).json({
      success: true,
      message: "File uploaded successfully",
      fileUrl,
      storagePath: destination,
    });
  } catch (error) {
    console.error("Firebase upload error:", error);
    return res.status(500).json({
      success: false,
      message: "Error uploading file to Firebase",
      error: error.message,
    });
  }
};

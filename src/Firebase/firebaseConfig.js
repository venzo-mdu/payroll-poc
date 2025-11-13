import admin from "firebase-admin";
import path from "path";
import dotenv from "dotenv";

dotenv.config();

const serviceAccountPath = path.resolve(process.env.SERVICE_ACCOUNT_KEY_PATH);

if (!admin.apps.length) {
  admin.initializeApp({
    credential: admin.credential.cert(serviceAccountPath),
    storageBucket: process.env.FIREBASE_STORAGE_BUCKET,
  });
}

const bucket = admin.storage().bucket();

export default bucket;

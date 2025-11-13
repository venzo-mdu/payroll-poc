import { Router } from "express";
import {  createUser } from "../controllers/user.controller.js";
import { uploadExcelToFirebase } from "../controllers/upload.controller.js";

const router = Router();

router.post("/", createUser);

router.post("/upload-excel", uploadExcelToFirebase);

// router.post("/single-sheet", createUserSingleSheet);

export default router;
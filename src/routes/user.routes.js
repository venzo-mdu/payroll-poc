import { Router } from "express";
import {  createUser } from "../controllers/user.controller.js";

const router = Router();

router.post("/", createUser);

// router.post("/single-sheet", createUserSingleSheet);

export default router;
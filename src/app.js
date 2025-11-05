import express from "express";
import dotenv from "dotenv";
import userRoutes from "./routes/user.routes.js";
import fileUpload from "express-fileupload";

dotenv.config();

const app = express();
app.use(express.json());

app.use(fileUpload());

// Routes
app.use("/api/users", userRoutes);

export default app;

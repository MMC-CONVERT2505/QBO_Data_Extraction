// index.js
import express from "express";
import dotenv from "dotenv";
import cors from "cors";

import qboRoutes from "./quickbooks.js";
import attachmentsRoutes from "./attachments.js";

dotenv.config();
const app = express();

app.set("trust proxy", true);
app.use(express.json());

// ✅ CORS
app.use(
  cors({
    origin: "*",
    methods: ["GET", "POST", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
  })
);
app.options("*", cors());

// ✅ AUTO PUBLIC_URL (always same host as request)
app.use((req, res, next) => {
  const proto = req.headers["x-forwarded-proto"] || req.protocol || "https";
  const host = req.get("host");
  global.PUBLIC_URL = `${proto}://${host}`;
  next();
});

// ✅ Routes
app.use("/", qboRoutes);
app.use("/", attachmentsRoutes);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Backend running at http://localhost:${PORT}`);
  console.log("✅ PUBLIC_URL will auto-detect from incoming requests.");
});

const express = require("express");
const mongoose = require("mongoose");
const cors = require("cors");
const os = require("os");
const path = require("path");
const fs = require("fs");

require("dotenv").config();

const app = express();
app.use(cors());
// SAP/MockTFL save payloads can include the full generated document state.
// Use a larger JSON body limit so "Save to Database" remains stable for larger studies.
app.use(express.json({ limit: "50mb" }));

// ==================== Database connection (same approach as AI-for-Excel) ====================
// Prefer Azure Cosmos DB for MongoDB (Mongo API), otherwise fall back to a standard MongoDB URI.
const MONGODB_URI = process.env.AZURE_COSMOS_URI || process.env.MONGODB_URI;
// If the URI does not include a database name (e.g. ends with "/?tls=true..."), Mongo defaults to "test".
// We want to align with SAP storage and use llxdocument by default.
const DEFAULT_DB_NAME = process.env.MONGODB_DB_NAME || process.env.DB_NAME || "llxdocument";

const dbSource = process.env.AZURE_COSMOS_URI ? "AZURE_COSMOS_URI" : "MONGODB_URI";
const dbType = process.env.AZURE_COSMOS_URI ? "Azure Cosmos DB (Mongo API)" : "MongoDB";

function maskMongoUri(uri) {
  if (!uri) return "";
  return uri.replace(/\/\/.*?:.*?@/, "//***:***@");
}

function stateToText(state) {
  return ({ 0: "disconnected", 1: "connected", 2: "connecting", 3: "disconnecting" }[state] || String(state));
}

function logDbStatus(stage) {
  console.log(`[DB] ${stage}`);
  console.log(`     type   : ${dbType}`);
  console.log(`     source : ${dbSource}`);
  console.log(`     state  : ${stateToText(mongoose.connection.readyState)}`);
  if (mongoose.connection.readyState === 1) {
    console.log(`     host   : ${mongoose.connection.host}`);
    console.log(`     dbName : ${mongoose.connection.name}`);
  }
}

console.log("[DB] URI set:", MONGODB_URI ? "yes" : "no");
console.log("[DB] URI masked:", maskMongoUri(MONGODB_URI));
logDbStatus("before connect");

// Azure Cosmos DB recommendation: retryWrites=false
const mongooseOptions = {
  retryWrites: false,
  maxIdleTimeMS: 120000,
  serverSelectionTimeoutMS: 30000,
  socketTimeoutMS: 360000,
  dbName: DEFAULT_DB_NAME
};

if (MONGODB_URI) {
  mongoose
    .connect(MONGODB_URI, mongooseOptions)
    .then(() => {
      console.log("Database connected.");
      logDbStatus("connected");
    })
    .catch((err) => {
      console.error("Database connection failed:", err.message);
      logDbStatus("after failure");
    });
} else {
  console.warn("Missing AZURE_COSMOS_URI or MONGODB_URI. Server will start, but database will NOT connect.");
}

// ==================== Minimal API (health checks only) ====================
app.get("/", (req, res) => {
  res.json({
    message: "AI-for-Word backend is running (minimal skeleton).",
    timestamp: new Date().toISOString(),
    database: {
      state: stateToText(mongoose.connection.readyState),
      type: dbType,
      source: dbSource,
      host: mongoose.connection.readyState === 1 ? mongoose.connection.host : null,
      dbName: mongoose.connection.readyState === 1 ? mongoose.connection.name : null
    }
  });
});

app.get("/api/test", (req, res) => {
  res.json({
    success: true,
    message: "AI-for-Word backend API ok",
    database: {
      state: stateToText(mongoose.connection.readyState),
      type: dbType,
      source: dbSource
    }
  });
});

// ==================== Routes ====================
app.use(require("./routes/references.routes"));
app.use(require("./routes/studies.routes"));
app.use(require("./routes/sap.routes"));
app.use(require("./routes/mocktfl.routes"));

// 404
app.use("*", (req, res) => {
  res.status(404).json({ success: false, message: "Not Found", path: req.originalUrl });
});

// Return JSON for request parsing failures as well, so the frontend never receives
// the default HTML error page for oversized or malformed payloads.
app.use((err, req, res, next) => {
  if (err?.type === "entity.too.large") {
    return res.status(413).json({
      success: false,
      message: "Request payload is too large. Please reduce the document size or increase the server limit."
    });
  }
  if (err instanceof SyntaxError && err?.status === 400 && "body" in err) {
    return res.status(400).json({ success: false, message: "Invalid JSON request body." });
  }
  return next(err);
});

// ==================== Start server (prefer HTTPS, fall back to HTTP) ====================
const PORT = Number(process.env.PORT || 4100);

function tryStartHttps() {
  try {
    const https = require("https");

    const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
    const keyPath = path.join(certDir, "localhost.key");
    const certPath = path.join(certDir, "localhost.crt");

    const key = fs.readFileSync(keyPath);
    const cert = fs.readFileSync(certPath);

    const httpsServer = https.createServer({ key, cert }, app);
    httpsServer.listen(PORT, () => {
      console.log("HTTPS started (office-addin-dev-certs).");
      console.log(`Listening on https://localhost:${PORT}`);
      console.log(`Database: ${dbType} (${dbSource})`);
      logDbStatus("server started (HTTPS)");
    });
    return true;
  } catch (err) {
    console.warn("HTTPS failed; falling back to HTTP:", err.message);
    return false;
  }
}

if (!tryStartHttps()) {
  app.listen(PORT, () => {
    console.log(`Listening on http://localhost:${PORT}`);
    console.log(`Database: ${dbType} (${dbSource})`);
    logDbStatus("server started (HTTP)");
  });
}

module.exports = app;

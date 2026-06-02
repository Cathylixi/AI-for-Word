const express = require("express");
const mongoose = require("mongoose");

const router = express.Router();

// Read SAP example entries from References.SAP_Example
router.get("/api/references/sap-example", async (req, res) => {
  try {
    console.log("[GET] /api/references/sap-example");
    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ success: false, message: "Database not connected" });
    }
    const db = mongoose.connection.client.db("References");
    const col = db.collection("SAP_Example");
    const docs = await col
      .find({}, { projection: { order: 1, number: 1, title: 1, _id: 0 } })
      .sort({ order: 1, number: 1 })
      .toArray();

    const entries = docs
      .filter((d) => d.number && d.title)
      .map((d) => ({
        order: d.order ?? null,
        number: String(d.number).trim(),
        title: String(d.title).trim()
      }));
    console.log(`[GET] /api/references/sap-example -> ${entries.length} entries`);
    res.json({ success: true, entries });
  } catch (err) {
    console.error("[GET] /api/references/sap-example error:", err);
    res.status(500).json({ success: false, message: "Failed to read SAP example", error: err.message });
  }
});

// Read MockTFL example entries from References.MockTFL_Example
router.get("/api/references/mocktfl-example", async (req, res) => {
  try {
    console.log("[GET] /api/references/mocktfl-example");
    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ success: false, message: "Database not connected" });
    }
    const db = mongoose.connection.client.db("References");
    const col = db.collection("MockTFL_Example");
    const docs = await col
      .find({}, { projection: { order: 1, type: 1, number: 1, title: 1, _id: 0 } })
      .sort({ order: 1, type: 1, number: 1 })
      .toArray();

    const entries = docs
      .filter((d) => d.type && d.number && d.title)
      .map((d) => ({
        order: d.order ?? null,
        type: String(d.type).trim().toUpperCase(),
        number: String(d.number).trim(),
        title: String(d.title).trim()
      }));
    console.log(`[GET] /api/references/mocktfl-example -> ${entries.length} entries`);
    res.json({ success: true, entries });
  } catch (err) {
    console.error("[GET] /api/references/mocktfl-example error:", err);
    res.status(500).json({ success: false, message: "Failed to read MockTFL example", error: err.message });
  }
});

module.exports = router;


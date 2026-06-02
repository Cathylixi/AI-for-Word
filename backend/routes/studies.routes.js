const express = require("express");
const mongoose = require("mongoose");

const router = express.Router();

// Read existing study numbers from llxdocument.existing study number
router.get("/api/studies/existing", async (req, res) => {
  try {
    console.log("[GET] /api/studies/existing");
    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ success: false, message: "Database not connected" });
    }
    const db = mongoose.connection.client.db("llxdocument");
    const col = db.collection("existing study number");
    const docs = await col.find({}, { projection: { studyNumber: 1, _id: 0 } }).toArray();
    const studies = docs
      .map((d) => String(d.studyNumber || "").trim())
      .filter(Boolean)
      .sort((a, b) => a.localeCompare(b, "en"));
    console.log(`[GET] /api/studies/existing -> ${studies.length} studies`);
    res.json({ success: true, studies });
  } catch (err) {
    console.error("[GET] /api/studies/existing error:", err);
    res
      .status(500)
      .json({ success: false, message: "Failed to read existing study numbers", error: err.message });
  }
});

module.exports = router;


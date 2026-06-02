const express = require("express");
const mongoose = require("mongoose");

const router = express.Router();

router.post("/api/sap/generate-section", async (req, res) => {
  try {
    console.log("[POST] /api/sap/generate-section body:", req.body);
    const { studyNumber, sapSectionNumber, refineInstruction } = req.body || {};
    if (!studyNumber || !sapSectionNumber) {
      return res.status(400).json({ success: false, message: "Missing studyNumber or sapSectionNumber" });
    }
    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ success: false, message: "Database not connected" });
    }
    const { generateSapSection } = require("../services/sap/sapGenerationService");
    const result = await generateSapSection({
      studyNumber: String(studyNumber).trim(),
      sapSectionNumber: String(sapSectionNumber).trim(),
      refineInstruction: refineInstruction ? String(refineInstruction).trim() : null,
      mongoose,
      openaiApiKey: process.env.OPENAI_API_KEY
    });
    console.log("[POST] /api/sap/generate-section matched:", result?.matchedProtocolSection);
    res.json({ success: true, ...result });
  } catch (err) {
    console.error("[POST] /api/sap/generate-section error:", err);
    res.status(500).json({ success: false, message: "Failed to generate SAP section", error: err.message });
  }
});

router.post("/api/sap/generate-titlepage-metadata", async (req, res) => {
  try {
    console.log("[POST] /api/sap/generate-titlepage-metadata body:", req.body);
    const { studyNumber } = req.body || {};
    if (!studyNumber) {
      return res.status(400).json({ success: false, message: "Missing studyNumber" });
    }
    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ success: false, message: "Database not connected" });
    }
    const { generateTitlePageMetadata } = require("../services/sap/titlePageMetadataService");
    const metadata = await generateTitlePageMetadata({
      studyNumber: String(studyNumber).trim(),
      mongoose,
      openaiApiKey: process.env.OPENAI_API_KEY
    });
    res.json({ success: true, metadata });
  } catch (err) {
    console.error("[POST] /api/sap/generate-titlepage-metadata error:", err);
    res
      .status(500)
      .json({ success: false, message: "Failed to generate title page metadata", error: err.message });
  }
});

router.post("/api/sap/generate-abbreviations", async (req, res) => {
  try {
    console.log("[POST] /api/sap/generate-abbreviations body:", req.body);
    const { studyNumber } = req.body || {};
    if (!studyNumber) {
      return res.status(400).json({ success: false, message: "Missing studyNumber" });
    }
    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ success: false, message: "Database not connected" });
    }
    const { generateAbbreviations } = require("../services/sap/abbreviationService");
    const items = await generateAbbreviations({
      studyNumber: String(studyNumber).trim(),
      mongoose,
      openaiApiKey: process.env.OPENAI_API_KEY
    });
    res.json({ success: true, items });
  } catch (err) {
    console.error("[POST] /api/sap/generate-abbreviations error:", err);
    res
      .status(500)
      .json({ success: false, message: "Failed to generate abbreviations", error: err.message });
  }
});

router.post("/api/sap/auto-select-sections", async (req, res) => {
  try {
    console.log("[POST] /api/sap/auto-select-sections body:", req.body);
    const { studyNumber, sapSections, threshold } = req.body || {};
    if (!studyNumber) {
      return res.status(400).json({ success: false, message: "Missing studyNumber" });
    }
    if (!Array.isArray(sapSections) || sapSections.length === 0) {
      return res.status(400).json({ success: false, message: "Missing sapSections" });
    }
    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ success: false, message: "Database not connected" });
    }
    const { autoSelectSapSections } = require("../services/sap/sapAutoSelectService");
    const recommended = await autoSelectSapSections({
      studyNumber: String(studyNumber).trim(),
      sapSections,
      threshold: typeof threshold === "number" ? threshold : undefined,
      mongoose
    });
    res.json({ success: true, recommended });
  } catch (err) {
    console.error("[POST] /api/sap/auto-select-sections error:", err);
    res.status(500).json({ success: false, message: "Failed to auto-select sections", error: err.message });
  }
});

router.post("/api/sap/extract-abbreviations-from-text", async (req, res) => {
  try {
    console.log("[POST] /api/sap/extract-abbreviations-from-text body:", req.body);
    const { text, existingTerms } = req.body || {};
    if (!text || !String(text).trim()) {
      return res.status(400).json({ success: false, message: "Missing text" });
    }
    const { extractAbbreviationsFromText } = require("../services/sap/abbreviationFromTextService");
    const items = await extractAbbreviationsFromText({
      text: String(text),
      existingTerms: Array.isArray(existingTerms) ? existingTerms : [],
      openaiApiKey: process.env.OPENAI_API_KEY
    });
    res.json({ success: true, items });
  } catch (err) {
    console.error("[POST] /api/sap/extract-abbreviations-from-text error:", err);
    res
      .status(500)
      .json({ success: false, message: "Failed to extract abbreviations", error: err.message });
  }
});

router.post("/api/sap/state", async (req, res) => {
  try {
    console.log("[POST] /api/sap/state body:", req.body);
    const { studyNumber, data } = req.body || {};
    if (!studyNumber) {
      return res.status(400).json({ success: false, message: "Missing studyNumber" });
    }
    if (!data || typeof data !== "object") {
      return res.status(400).json({ success: false, message: "Missing data payload" });
    }
    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ success: false, message: "Database not connected" });
    }

    const db = mongoose.connection.client.db("llxdocument");
    const col = db.collection("studies");

    // IMPORTANT:
    // We enrich `data.entries` with cached/new title embeddings on the backend so:
    // - OpenAI keys never touch the frontend
    // - We avoid re-embedding if the title was already embedded and saved
    // - We preserve the user's assumption that titles do not change
    let enrichedData = data;
    try {
      const existing = await col.findOne(
        { studyNumber: String(studyNumber).trim(), chunkType: "SAP" },
        { projection: { data: 1, _id: 0 } }
      );
      const { ensureSapEntryTitleEmbeddings } = require("../services/sap/sapTitleEmbeddingService");
      enrichedData = await ensureSapEntryTitleEmbeddings({
        studyNumber: String(studyNumber).trim(),
        incomingState: data,
        existingState: existing?.data || null,
        mongoose,
        openaiApiKey: process.env.OPENAI_API_KEY
      });
    } catch (e) {
      // If we cannot generate embeddings (e.g., missing OPENAI_API_KEY), we fail the save
      // because the user explicitly wants caching behavior and stable embeddings.
      return res.status(500).json({
        success: false,
        message: "Failed to save SAP state (embedding step failed)",
        error: e?.message || String(e)
      });
    }

    const result = await col.updateOne(
      { studyNumber, chunkType: "SAP" },
      {
        $set: {
          data: enrichedData,
          updatedAt: new Date()
        },
        $setOnInsert: {
          createdAt: new Date(),
          formKey: null,
          formOrder: null
        }
      },
      { upsert: true }
    );

    res.json({ success: true, result });
  } catch (err) {
    console.error("[POST] /api/sap/state error:", err);
    res.status(500).json({ success: false, message: "Failed to save SAP state", error: err.message });
  }
});

router.get("/api/sap/state", async (req, res) => {
  try {
    const { studyNumber } = req.query || {};
    console.log(`[GET] /api/sap/state?studyNumber=${studyNumber}`);
    if (!studyNumber) {
      return res.status(400).json({ success: false, message: "Missing studyNumber" });
    }
    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ success: false, message: "Database not connected" });
    }

    const db = mongoose.connection.client.db("llxdocument");
    const col = db.collection("studies");

    const doc = await col.findOne({ studyNumber, chunkType: "SAP" }, { projection: { data: 1, _id: 0 } });

    res.json({ success: true, data: doc ? doc.data : null });
  } catch (err) {
    console.error("[GET] /api/sap/state error:", err);
    res.status(500).json({ success: false, message: "Failed to load SAP state", error: err.message });
  }
});

module.exports = router;

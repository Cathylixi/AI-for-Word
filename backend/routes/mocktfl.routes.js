const express = require("express");
const mongoose = require("mongoose");

const router = express.Router();

function isAllowedMockTflAnalysisTarget(type, number) {
  const t = String(type || "").trim().toUpperCase();
  const n = String(number || "").trim();
  const allowed = new Set(["14.1.1.2", "14.1.3.1", "14.2.1.1"]);
  return t === "TABLE" && allowed.has(n);
}

function isAllowedMockTflFigureTarget(type, number) {
  const t = String(type || "").trim().toUpperCase();
  const n = String(number || "").trim();
  const allowed = new Set(["14.2.1.1"]);
  return t === "FIGURE" && allowed.has(n);
}

// Auto-select MockTFL sections based on protocol embeddings
router.post("/api/mocktfl/auto-select-sections", async (req, res) => {
  try {
    console.log("[POST] /api/mocktfl/auto-select-sections body:", req.body);
    const { studyNumber, tflSections, threshold } = req.body || {};
    if (!studyNumber) {
      return res.status(400).json({ success: false, message: "Missing studyNumber" });
    }
    if (!Array.isArray(tflSections) || tflSections.length === 0) {
      return res.status(400).json({ success: false, message: "Missing tflSections" });
    }
    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ success: false, message: "Database not connected" });
    }
    const { autoSelectMockTflSections } = require("../services/mocktfl/mockTflAutoSelectService");
    const recommended = await autoSelectMockTflSections({
      studyNumber: String(studyNumber).trim(),
      tflSections,
      threshold: typeof threshold === "number" ? threshold : undefined,
      mongoose
    });
    res.json({ success: true, recommended });
  } catch (err) {
    console.error("[POST] /api/mocktfl/auto-select-sections error:", err);
    res.status(500).json({ success: false, message: "Failed to auto-select MockTFL sections", error: err.message });
  }
});

router.post("/api/mocktfl/generate-figure", async (req, res) => {
  try {
    console.log("[POST] /api/mocktfl/generate-figure body:", {
      studyNumber: req.body?.studyNumber,
      type: req.body?.type,
      number: req.body?.number,
      rCodeLength: String(req.body?.rCodeText || "").length
    });
    const { studyNumber, type, number, rCodeText } = req.body || {};
    if (!studyNumber) {
      return res.status(400).json({ success: false, message: "Missing studyNumber" });
    }
    if (!rCodeText || !String(rCodeText).trim()) {
      return res.status(400).json({ success: false, message: "Missing R code" });
    }
    if (!isAllowedMockTflFigureTarget(type, number)) {
      return res.status(400).json({
        success: false,
        message: "Figure generation is only enabled for FIGURE 14.2.1.1 in the current prototype."
      });
    }

    const { generateMockTflFigureFromRCode } = require("../services/mocktfl/mockTflFigureService");
    const result = await generateMockTflFigureFromRCode({
      rCodeText: String(rCodeText || "")
    });

    res.json({ success: true, ...result });
  } catch (err) {
    console.error("[POST] /api/mocktfl/generate-figure error:", err);
    res.status(500).json({
      success: false,
      message: err?.message || "Failed to generate MockTFL figure",
      error: err.message
    });
  }
});

router.post("/api/mocktfl/auto-generate-figure", async (req, res) => {
  try {
    console.log("[POST] /api/mocktfl/auto-generate-figure body:", {
      studyNumber: req.body?.studyNumber,
      type: req.body?.type,
      number: req.body?.number,
      pureTitleLength: String(req.body?.pureTitle || "").length,
      drug1: req.body?.drug1,
      drug2: req.body?.drug2
    });
    const {
      studyNumber,
      type,
      number,
      pureTitle,
      drug1,
      drug2,
      figureType,
      correspondingTableType,
      correspondingTableNumber
    } = req.body || {};
    if (!studyNumber) {
      return res.status(400).json({ success: false, message: "Missing studyNumber" });
    }
    if (!isAllowedMockTflFigureTarget(type, number)) {
      return res.status(400).json({
        success: false,
        message: "Auto figure generation is only enabled for FIGURE 14.2.1.1 in the current prototype."
      });
    }

    const { autoGenerateMockTflFigure } = require("../services/mocktfl/mockTflFigureAutoService");
    const result = await autoGenerateMockTflFigure({
      studyNumber: String(studyNumber || "").trim(),
      type: String(type || "").trim(),
      number: String(number || "").trim(),
      pureTitle: String(pureTitle || "").trim(),
      drug1: String(drug1 || "").trim(),
      drug2: String(drug2 || "").trim(),
      figureType: String(figureType || "").trim(),
      correspondingTableType: String(correspondingTableType || "").trim(),
      correspondingTableNumber: String(correspondingTableNumber || "").trim(),
      mongoose,
      openaiApiKey: process.env.OPENAI_API_KEY
    });

    res.json({ success: true, ...result });
  } catch (err) {
    console.error("[POST] /api/mocktfl/auto-generate-figure error:", err);
    res.status(500).json({
      success: false,
      message: err?.message || "Failed to auto-generate MockTFL figure",
      error: err.message
    });
  }
});

router.post("/api/mocktfl/find-corresponding-table", async (req, res) => {
  try {
    console.log("[POST] /api/mocktfl/find-corresponding-table body:", {
      studyNumber: req.body?.studyNumber,
      type: req.body?.type,
      number: req.body?.number,
      pureTitleLength: String(req.body?.pureTitle || "").length
    });
    const { studyNumber, type, number, pureTitle } = req.body || {};
    if (!studyNumber) {
      return res.status(400).json({ success: false, message: "Missing studyNumber" });
    }
    if (!isAllowedMockTflFigureTarget(type, number)) {
      return res.status(400).json({
        success: false,
        message: "Find corresponding table is only enabled for FIGURE 14.2.1.1 in the current prototype."
      });
    }
    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ success: false, message: "Database not connected" });
    }

    const { findCorrespondingTable } = require("../services/mocktfl/mockTflCorrespondingTableService");
    const result = await findCorrespondingTable({
      studyNumber: String(studyNumber || "").trim(),
      type: String(type || "").trim(),
      number: String(number || "").trim(),
      pureTitle: String(pureTitle || "").trim(),
      mongoose,
      openaiApiKey: process.env.OPENAI_API_KEY
    });

    res.json({ success: true, ...result });
  } catch (err) {
    console.error("[POST] /api/mocktfl/find-corresponding-table error:", err);
    res.status(500).json({
      success: false,
      message: err?.message || "Failed to find corresponding table",
      error: err.message
    });
  }
});

router.post("/api/mocktfl/define-figure-type", async (req, res) => {
  try {
    console.log("[POST] /api/mocktfl/define-figure-type body:", {
      studyNumber: req.body?.studyNumber,
      type: req.body?.type,
      number: req.body?.number,
      pureTitleLength: String(req.body?.pureTitle || "").length,
      correspondingTableNumber: req.body?.correspondingTableNumber
    });
    const {
      studyNumber,
      type,
      number,
      pureTitle,
      correspondingTableType,
      correspondingTableNumber
    } = req.body || {};
    if (!studyNumber) {
      return res.status(400).json({ success: false, message: "Missing studyNumber" });
    }
    if (!isAllowedMockTflFigureTarget(type, number)) {
      return res.status(400).json({
        success: false,
        message: "Defining figure type is only enabled for FIGURE 14.2.1.1 in the current prototype."
      });
    }
    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ success: false, message: "Database not connected" });
    }

    const { defineMockTflFigureType } = require("../services/mocktfl/mockTflFigureTypeService");
    const result = await defineMockTflFigureType({
      studyNumber: String(studyNumber || "").trim(),
      pureTitle: String(pureTitle || "").trim(),
      tableType: String(correspondingTableType || "").trim(),
      tableNumber: String(correspondingTableNumber || "").trim(),
      mongoose,
      openaiApiKey: process.env.OPENAI_API_KEY
    });

    res.json({ success: true, ...result });
  } catch (err) {
    console.error("[POST] /api/mocktfl/define-figure-type error:", err);
    res.status(500).json({
      success: false,
      message: err?.message || "Failed to define figure type",
      error: err.message
    });
  }
});

/**
 * Analyze a specific MockTFL item (prototype allowlist):
 * - Embed the (pure) title
 * - Match against CRF form title embeddings (chunkType=crf_crfFormList)
 * - Return all questions from the best matched CRF form
 */
router.post("/api/mocktfl/analyze-crf-questions", async (req, res) => {
  try {
    console.log("[POST] /api/mocktfl/analyze-crf-questions body:", req.body);
    const { studyNumber, type, number, pureTitle } = req.body || {};

    if (!studyNumber) return res.status(400).json({ success: false, message: "Missing studyNumber" });
    if (!pureTitle) return res.status(400).json({ success: false, message: "Missing pureTitle" });
    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ success: false, message: "Database not connected" });
    }

    // Feature gate:
    // We only support a small allowlist of MockTFL work items for the current prototype.
    // IMPORTANT: This is intentionally strict so we don't accidentally enable analysis for all tables yet.
    if (!isAllowedMockTflAnalysisTarget(type, number)) {
      return res.status(400).json({
        success: false,
        message: "Start analysis is only enabled for TABLE 14.1.1.2, TABLE 14.1.3.1 and TABLE 14.2.1.1 in the current prototype."
      });
    }

    const { analyzeMockTflCrfQuestions } = require("../services/mocktfl/mockTflCrfQuestionService");
    const result = await analyzeMockTflCrfQuestions({
      studyNumber: String(studyNumber).trim(),
      pureTitle: String(pureTitle).trim(),
      mongoose,
      openaiApiKey: process.env.OPENAI_API_KEY
    });

    res.json({ success: true, ...result });
  } catch (err) {
    console.error("[POST] /api/mocktfl/analyze-crf-questions error:", err);
    res.status(500).json({
      success: false,
      message: "Failed to analyze MockTFL CRF questions",
      error: err.message
    });
  }
});

/**
 * Step 1 Start Analysis endpoint for MockTFL:
 * - Embed pure title once
 * - SAP branch: match saved SAP section by title embedding
 * - SAP branch: ask AI to extract editable statistical variables
 */
router.post("/api/mocktfl/start-analysis", async (req, res) => {
  try {
    console.log("[POST] /api/mocktfl/start-analysis body:", req.body);
    const { studyNumber, type, number, pureTitle } = req.body || {};

    if (!studyNumber) return res.status(400).json({ success: false, message: "Missing studyNumber" });
    if (!pureTitle) return res.status(400).json({ success: false, message: "Missing pureTitle" });
    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ success: false, message: "Database not connected" });
    }

    // Feature gate (same as current prototype behavior).
    if (!isAllowedMockTflAnalysisTarget(type, number)) {
      return res.status(400).json({
        success: false,
        message: "Start analysis is only enabled for TABLE 14.1.1.2, TABLE 14.1.3.1 and TABLE 14.2.1.1 in the current prototype."
      });
    }

    const { startMockTflAnalysis } = require("../services/mocktfl/mockTflStartAnalysisService");
    const result = await startMockTflAnalysis({
      studyNumber: String(studyNumber).trim(),
      pureTitle: String(pureTitle).trim(),
      mongoose,
      openaiApiKey: process.env.OPENAI_API_KEY
    });

    res.json({ success: true, ...result });
  } catch (err) {
    console.error("[POST] /api/mocktfl/start-analysis error:", err);
    res.status(500).json({
      success: false,
      message: "Failed to run MockTFL start analysis",
      error: err.message
    });
  }
});

/**
 * Step 2 endpoint for MockTFL:
 * - Accept the user-confirmed statistical variables
 * - Map them to the best matching CRF questions and answers
 */
router.post("/api/mocktfl/map-variables-to-crf", async (req, res) => {
  try {
    console.log("[POST] /api/mocktfl/map-variables-to-crf body:", req.body);
    const { studyNumber, type, number, pureTitle, variables, variablesText, savedCrfFormName } = req.body || {};

    if (!studyNumber) return res.status(400).json({ success: false, message: "Missing studyNumber" });
    if (!pureTitle) return res.status(400).json({ success: false, message: "Missing pureTitle" });
    if (!Array.isArray(variables) && !String(variablesText || "").trim()) {
      return res.status(400).json({ success: false, message: "Missing confirmed variables" });
    }
    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ success: false, message: "Database not connected" });
    }

    if (!isAllowedMockTflAnalysisTarget(type, number)) {
      return res.status(400).json({
        success: false,
        message: "Start analysis is only enabled for TABLE 14.1.1.2, TABLE 14.1.3.1 and TABLE 14.2.1.1 in the current prototype."
      });
    }

    const { mapMockTflVariablesToCrf } = require("../services/mocktfl/mockTflStartAnalysisService");
    const result = await mapMockTflVariablesToCrf({
      studyNumber: String(studyNumber).trim(),
      pureTitle: String(pureTitle).trim(),
      variables,
      variablesText: String(variablesText || ""),
      savedCrfFormName: String(savedCrfFormName || "").trim(),
      mongoose,
      openaiApiKey: process.env.OPENAI_API_KEY
    });

    res.json({ success: true, ...result });
  } catch (err) {
    console.error("[POST] /api/mocktfl/map-variables-to-crf error:", err);
    res.status(500).json({
      success: false,
      message: "Failed to map confirmed variables to CRF questions",
      error: err.message
    });
  }
});

// Save MockTFL state
/**
 * Extract drug comparison from SAP Introduction
 */
router.post("/api/mocktfl/drug-comparison", async (req, res) => {
  try {
    console.log("[POST] /api/mocktfl/drug-comparison body:", req.body);
    const { studyNumber } = req.body || {};

    if (!studyNumber) return res.status(400).json({ success: false, message: "Missing studyNumber" });
    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ success: false, message: "Database not connected" });
    }

    const { extractDrugComparisonFromSap } = require("../services/mocktfl/mockTflStartAnalysisService");
    const result = await extractDrugComparisonFromSap({
      studyNumber: String(studyNumber).trim(),
      mongoose,
      openaiApiKey: process.env.OPENAI_API_KEY
    });

    res.json({ success: true, ...result });
  } catch (err) {
    console.error("[POST] /api/mocktfl/drug-comparison error:", err);
    res.status(500).json({
      success: false,
      message: "Failed to extract drug comparison from SAP",
      error: err.message
    });
  }
});

router.post("/api/mocktfl/state", async (req, res) => {
  try {
    console.log("[POST] /api/mocktfl/state body:", req.body);
    const { studyNumber, data } = req.body || {};
    if (!studyNumber) {
      return res.status(400).json({ success: false, message: "Missing studyNumber" });
    }
    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ success: false, message: "Database not connected" });
    }

    const db = mongoose.connection.client.db("llxdocument");
    const col = db.collection("studies");

    const result = await col.updateOne(
      { studyNumber, chunkType: "MockTFL" },
      {
        $set: {
          data,
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
    console.error("[POST] /api/mocktfl/state error:", err);
    res.status(500).json({ success: false, message: "Failed to save MockTFL state", error: err.message });
  }
});

// Load MockTFL state
router.get("/api/mocktfl/state", async (req, res) => {
  try {
    const { studyNumber } = req.query || {};
    console.log(`[GET] /api/mocktfl/state?studyNumber=${studyNumber}`);
    if (!studyNumber) {
      return res.status(400).json({ success: false, message: "Missing studyNumber" });
    }
    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ success: false, message: "Database not connected" });
    }

    const db = mongoose.connection.client.db("llxdocument");
    const col = db.collection("studies");

    const doc = await col.findOne({ studyNumber, chunkType: "MockTFL" }, { projection: { data: 1, _id: 0 } });

    res.json({ success: true, data: doc ? doc.data : null });
  } catch (err) {
    console.error("[GET] /api/mocktfl/state error:", err);
    res.status(500).json({ success: false, message: "Failed to load MockTFL state", error: err.message });
  }
});

module.exports = router;


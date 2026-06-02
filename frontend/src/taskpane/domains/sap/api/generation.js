import { apiPost } from "../../../shared/api/client";

async function generateSapSection({ studyNumber, sapSectionNumber, refineInstruction }) {
  const data = await apiPost("/api/sap/generate-section", {
    studyNumber,
    sapSectionNumber,
    refineInstruction
  });
  return data || {};
}

async function generateTitlePageMetadata({ studyNumber }) {
  const data = await apiPost("/api/sap/generate-titlepage-metadata", {
    studyNumber
  });
  return data?.metadata || {};
}

async function generateAbbreviations({ studyNumber }) {
  const data = await apiPost("/api/sap/generate-abbreviations", {
    studyNumber
  });
  return Array.isArray(data?.items) ? data.items : [];
}

async function extractAbbreviationsFromText({ text, existingTerms }) {
  const data = await apiPost("/api/sap/extract-abbreviations-from-text", {
    text,
    existingTerms
  });
  return Array.isArray(data?.items) ? data.items : [];
}

async function autoSelectSapSections({ studyNumber, sapSections, threshold }) {
  const data = await apiPost("/api/sap/auto-select-sections", {
    studyNumber,
    sapSections,
    threshold
  });
  return Array.isArray(data?.recommended) ? data.recommended : [];
}

export {
  generateSapSection,
  generateTitlePageMetadata,
  generateAbbreviations,
  extractAbbreviationsFromText,
  autoSelectSapSections
};

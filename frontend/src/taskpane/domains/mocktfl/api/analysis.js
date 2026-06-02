import { apiPost } from "../../../shared/api/client";

/**
 * Ask backend to:
 * - Embed the pure MockTFL title (without "TABLE 14.x" prefix)
 * - Match against CRF form title embeddings for the given study
 * - Return all question strings from the best matched CRF form
 *
 * IMPORTANT:
 * - Backend currently enforces a strict feature gate (prototype allowlist).
 */
async function analyzeMockTflCrfQuestions({ studyNumber, type, number, pureTitle }) {
  return await apiPost("/api/mocktfl/analyze-crf-questions", {
    studyNumber,
    type,
    number,
    pureTitle
  });
}

export { analyzeMockTflCrfQuestions };


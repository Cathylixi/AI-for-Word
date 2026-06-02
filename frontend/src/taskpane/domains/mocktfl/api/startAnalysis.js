import { apiPost } from "../../../shared/api/client";

/**
 * Step 1 of MockTFL analysis (prototype allowlist):
 * - Backend will embed pureTitle once
 * - SAP branch: match best saved SAP section
 * - GPT branch: return editable suggested statistical variables
 */
async function startMockTflAnalysis({ studyNumber, type, number, pureTitle }) {
  return await apiPost("/api/mocktfl/start-analysis", {
    studyNumber,
    type,
    number,
    pureTitle
  });
}

export { startMockTflAnalysis };


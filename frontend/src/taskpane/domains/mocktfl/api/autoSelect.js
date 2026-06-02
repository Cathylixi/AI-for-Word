import { apiPost } from "../../../shared/api/client";

async function autoSelectMockTflSections({ studyNumber, tflSections, threshold }) {
  const data = await apiPost("/api/mocktfl/auto-select-sections", {
    studyNumber,
    tflSections,
    threshold
  });
  return Array.isArray(data?.recommended) ? data.recommended : [];
}

export { autoSelectMockTflSections };

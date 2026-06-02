import { apiPost } from "../../../shared/api/client";

async function generateMockTflFigure({ studyNumber, type, number, rCodeText }) {
  return await apiPost("/api/mocktfl/generate-figure", {
    studyNumber,
    type,
    number,
    rCodeText
  });
}

export { generateMockTflFigure };

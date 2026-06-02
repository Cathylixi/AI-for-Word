import { apiPost } from "../../../shared/api/client";

async function autoGenerateMockTflFigure({
  studyNumber,
  type,
  number,
  pureTitle,
  drug1,
  drug2,
  figureType,
  correspondingTableType,
  correspondingTableNumber
}) {
  return await apiPost("/api/mocktfl/auto-generate-figure", {
    studyNumber,
    type,
    number,
    pureTitle,
    drug1,
    drug2,
    figureType,
    correspondingTableType,
    correspondingTableNumber
  });
}

export { autoGenerateMockTflFigure };


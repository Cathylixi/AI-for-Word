import { apiPost } from "../../../shared/api/client";

async function defineMockTflFigureType({
  studyNumber,
  type,
  number,
  pureTitle,
  correspondingTableType,
  correspondingTableNumber
}) {
  return await apiPost("/api/mocktfl/define-figure-type", {
    studyNumber,
    type,
    number,
    pureTitle,
    correspondingTableType,
    correspondingTableNumber
  });
}

export { defineMockTflFigureType };

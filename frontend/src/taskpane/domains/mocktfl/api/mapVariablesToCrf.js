import { apiPost } from "../../../shared/api/client";

/**
 * Step 2 of MockTFL analysis:
 * - Send the user-confirmed variables back to the backend
 * - Ask the backend to map them to CRF questions and answers
 */
async function mapMockTflVariablesToCrf({
  studyNumber,
  type,
  number,
  pureTitle,
  variables,
  variablesText,
  savedCrfFormName
}) {
  return await apiPost("/api/mocktfl/map-variables-to-crf", {
    studyNumber,
    type,
    number,
    pureTitle,
    variables,
    variablesText,
    savedCrfFormName
  });
}

export { mapMockTflVariablesToCrf };

export * from "../../domains/mocktfl/api/state";

import { apiGet, apiPost } from "../client";

async function loadMockTflState({ studyNumber }) {
  return await apiGet(`/api/mocktfl/state?studyNumber=${encodeURIComponent(studyNumber)}`);
}

async function saveMockTflState({ studyNumber, state }) {
  return await apiPost("/api/mocktfl/state", {
    studyNumber,
    data: state
  });
}

export { loadMockTflState, saveMockTflState };

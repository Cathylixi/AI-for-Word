import { apiGet, apiPost } from "../../../shared/api/client";

async function loadSapState({ studyNumber }) {
  // Returns { success: true, data: { ... } } or { success: true, data: null }
  return await apiGet(`/api/sap/state?studyNumber=${encodeURIComponent(studyNumber)}`);
}

async function saveSapState({ studyNumber, state }) {
  // Returns { success: true }
  return await apiPost("/api/sap/state", {
    studyNumber,
    data: state
  });
}

export { loadSapState, saveSapState };

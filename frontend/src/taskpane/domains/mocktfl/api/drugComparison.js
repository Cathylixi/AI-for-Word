import { apiPost } from "../../../shared/api/client";

export async function extractDrugComparison({ studyNumber }) {
  if (!studyNumber) throw new Error("Missing studyNumber for drug comparison extraction");
  return await apiPost("/api/mocktfl/drug-comparison", { studyNumber });
}

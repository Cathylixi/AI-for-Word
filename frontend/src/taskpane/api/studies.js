import { apiGet } from "./client";

async function getExistingStudyNumbers() {
  const data = await apiGet("/api/studies/existing");
  return Array.isArray(data?.studies) ? data.studies : [];
}

export { getExistingStudyNumbers };

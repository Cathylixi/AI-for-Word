import { apiGet } from "../../../shared/api/client";

async function getMockTflExampleEntries() {
  const data = await apiGet("/api/references/mocktfl-example");
  return Array.isArray(data?.entries) ? data.entries : [];
}

export { getMockTflExampleEntries };

export * from "../../domains/mocktfl/api/references";

import { apiGet } from "../client";

async function getMockTflExampleEntries() {
  const data = await apiGet("/api/references/mocktfl-example");
  return Array.isArray(data?.entries) ? data.entries : [];
}

export { getMockTflExampleEntries };

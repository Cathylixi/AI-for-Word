import { apiGet } from "../../../shared/api/client";

async function getSapExampleEntries() {
  const data = await apiGet("/api/references/sap-example");
  return Array.isArray(data?.entries) ? data.entries : [];
}

export { getSapExampleEntries };

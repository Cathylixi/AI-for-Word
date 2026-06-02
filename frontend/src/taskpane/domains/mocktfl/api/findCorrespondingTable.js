import { apiPost } from "../../../shared/api/client";

async function findCorrespondingTable({ studyNumber, type, number, pureTitle }) {
  return await apiPost("/api/mocktfl/find-corresponding-table", {
    studyNumber,
    type,
    number,
    pureTitle
  });
}

export { findCorrespondingTable };

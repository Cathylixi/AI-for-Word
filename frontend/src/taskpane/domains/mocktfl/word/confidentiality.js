/* global Word */

function insertConfidentialityStatement(body) {
  const title = body.insertParagraph("Confidentiality Statement", Word.InsertLocation.end);
  title.alignment = "Centered";
  title.font.set({ name: "Times New Roman", size: 12, bold: true });

  body.insertParagraph("", Word.InsertLocation.end);

  const p = body.insertParagraph(
    "This document is confidential. It contains proprietary information belonging to Anthos Therapeutics, Inc. Any viewing or disclosure of such information that is not authorized in writing by Anthos Therapeutics is strictly prohibited. Such information may be solely for the purpose of reviewing or performing this study.",
    Word.InsertLocation.end
  );
  p.alignment = "Left";
  p.font.set({ name: "Times New Roman", size: 12, bold: false });
}

export { insertConfidentialityStatement };


export * from "../../domains/mocktfl/word/titlePage";

/* global Word */

function insertMockTflTitlePage(body, { protocolNo = "Protocol no. ANT-007" } = {}) {
  // This is a minimal first version mirroring the screenshot structure.
  // Header-ish lines (we insert as normal paragraphs; true Word headers can be added later).
  const h1 = body.insertParagraph("Anthos Therapeutics, Inc., A Novartis Company", Word.InsertLocation.end);
  h1.font.set({ name: "Times New Roman", size: 10, bold: false });
  const h2 = body.insertParagraph(protocolNo, Word.InsertLocation.end);
  h2.font.set({ name: "Times New Roman", size: 10, bold: false });

  body.insertParagraph("", Word.InsertLocation.end);
  body.insertParagraph("", Word.InsertLocation.end);

  const title = body.insertParagraph("Clinical Study Report Analysis:", Word.InsertLocation.end);
  title.alignment = "Centered";
  title.font.set({ name: "Times New Roman", size: 28, bold: true });

  const subtitle = body.insertParagraph("Mock Tables, Figures & Listings", Word.InsertLocation.end);
  subtitle.alignment = "Centered";
  subtitle.font.set({ name: "Times New Roman", size: 28, bold: true });

  body.insertParagraph("", Word.InsertLocation.end);
  body.insertParagraph("", Word.InsertLocation.end);

  const studyLine = body.insertParagraph(
    "A MULTICENTER, RANDOMIZED, OPEN-LABEL, BLINDED ENDPOINT EVALUATION, PHASE 3 STUDY COMPARING THE EFFECT OF ABELACIMAB\nRELATIVE TO APIXABAN ON VENOUS THROMBOEMBOLISM (VTE) RECURRENCE AND BLEEDING IN PATIENTS WITH CANCER ASSOCIATED VTE\n(ASTER)",
    Word.InsertLocation.end
  );
  studyLine.alignment = "Centered";
  studyLine.font.set({ name: "Times New Roman", size: 12, bold: true });

  body.insertParagraph("", Word.InsertLocation.end);
  body.insertParagraph("", Word.InsertLocation.end);

  const version = body.insertParagraph("Version: 3", Word.InsertLocation.end);
  version.alignment = "Centered";
  version.font.set({ name: "Times New Roman", size: 12, bold: false });

  const date = body.insertParagraph("Date: 30JAN2026", Word.InsertLocation.end);
  date.alignment = "Centered";
  date.font.set({ name: "Times New Roman", size: 12, bold: false });

  body.insertParagraph("", Word.InsertLocation.end);
  body.insertParagraph("", Word.InsertLocation.end);

  const proto = body.insertParagraph("Protocol No: ANT-007", Word.InsertLocation.end);
  proto.alignment = "Centered";
  proto.font.set({ name: "Times New Roman", size: 12, bold: false });

  body.insertParagraph("", Word.InsertLocation.end);
  body.insertParagraph("", Word.InsertLocation.end);

  const submitted = body.insertParagraph("Submitted to:", Word.InsertLocation.end);
  submitted.alignment = "Centered";
  submitted.font.set({ name: "Times New Roman", size: 12, bold: false });

  const company = body.insertParagraph("Anthos Therapeutics, Inc., A Novartis Company", Word.InsertLocation.end);
  company.alignment = "Centered";
  company.font.set({ name: "Times New Roman", size: 12, bold: true });

  body.insertParagraph("", Word.InsertLocation.end);
  body.insertParagraph("", Word.InsertLocation.end);

  const prepared = body.insertParagraph("Prepared by:", Word.InsertLocation.end);
  prepared.alignment = "Centered";
  prepared.font.set({ name: "Times New Roman", size: 12, bold: false });

  const vendor = body.insertParagraph("LLX Solutions, LLC", Word.InsertLocation.end);
  vendor.alignment = "Centered";
  vendor.font.set({ name: "Times New Roman", size: 12, bold: true });
}

export { insertMockTflTitlePage };


/* global Word */

import { insertTitlePage } from "./titlePage";
import { insertRevisionHistory } from "./revisionHistory";
import { insertTableOfContents } from "./toc";
import { insertListOfAbbreviationsPlaceholder } from "./listOfAbbreviations";
import { TAG_PREFIX_BODY, TAG_PREFIX_TITLE, TAG_PREFIX_META } from "./constants";

async function clearSapTemplate() {
  await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items");
    await context.sync();

    // Load tags safely (avoid collection load path issues)
    ccs.items.forEach((cc) => cc.load("tag"));
    await context.sync();

    const toDelete = ccs.items.filter((cc) => {
      const tag = cc.tag || "";
      return (
        tag.startsWith(TAG_PREFIX_TITLE) ||
        tag.startsWith(TAG_PREFIX_BODY) ||
        tag.startsWith(TAG_PREFIX_META)
      );
    });

    // Delete in reverse order to reduce range conflicts
    for (let i = toDelete.length - 1; i >= 0; i--) {
      toDelete[i].delete(false);
    }
    await context.sync();
  });
}

function applyHeadingStyle(paragraph, level) {
  try {
    if (Word?.StyleBuiltIn) {
      const map = {
        1: Word.StyleBuiltIn.heading1,
        2: Word.StyleBuiltIn.heading2,
        3: Word.StyleBuiltIn.heading3
      };
      paragraph.styleBuiltIn = map[level] || Word.StyleBuiltIn.heading1;
      return;
    }
  } catch (e) {}
  try {
    paragraph.style = `Heading ${level}`;
  } catch (e) {}
}

// Generate the front matter pages up to and including the Table of Contents.
// This function intentionally does NOT generate the Abbreviations page or main sections,
// so the controller can strictly follow "page order" (TOC first, then fetch abbreviations, then insert).
async function generateSapFrontMatter({ studyNumber, entries }) {
  await Word.run(async (context) => {
    const body = context.document.body;

    // Title page / header section (includes Signature Page within the title page)
    insertTitlePage(body, studyNumber);

    // New page for Revision History
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
    insertRevisionHistory(body);

    // New page for Table of Contents
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
    insertTableOfContents(body);

    await context.sync();
  });
}

// Continue document generation AFTER the Table of Contents.
// Inserts List of Abbreviations (placeholder) and then the main SAP sections (new page).
async function generateSapAfterToc({ studyNumber, entries }) {
  await Word.run(async (context) => {
    const body = context.document.body;

    // New page for List of Abbreviations (after TOC)
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
    insertListOfAbbreviationsPlaceholder(body);

    // New page for main SAP sections
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

    entries.forEach((entry) => {
      const number = entry.number || ""; // original SAP number for internal tags/mapping
      const displayNumber = entry.displayNumber || number; // renumbered display number for headings/TOC
      const title = entry.title || "";
      const label = `${displayNumber} ${title}`.trim();

      // Title line wrapped in a content control
      const titlePara = body.insertParagraph(label, Word.InsertLocation.end);

      // Determine heading level and font size based on dot count in number (hard rule per user request)
      let headingLevel = 1;
      let fontSize = 14;
      if (displayNumber) {
        const dots = (displayNumber.match(/\./g) || []).length;
        if (dots === 1) {
          headingLevel = 2;
          fontSize = 13;
        }
        if (dots === 2) {
          headingLevel = 3;
          fontSize = 12;
        }
      }

      const titleRange = titlePara.getRange();
      applyHeadingStyle(titlePara, headingLevel);
      titleRange.font.set({
        name: "Arial",
        bold: true,
        size: fontSize,
        color: "#000000"
      });

      const titleCc = titleRange.insertContentControl();
      titleCc.tag = `${TAG_PREFIX_TITLE}${studyNumber}:${number}`;
      titleCc.title = label;
      titleCc.appearance = "BoundingBox";

      // Body content control
      const bodyPara = body.insertParagraph("", Word.InsertLocation.end);
      const bodyRange = bodyPara.getRange();
      bodyRange.font.set({
        name: "Times New Roman",
        size: 12,
        bold: false
      });

      const bodyCc = bodyRange.insertContentControl();
      bodyCc.tag = `${TAG_PREFIX_BODY}${studyNumber}:${number}`;
      bodyCc.title = label;
      bodyCc.appearance = "BoundingBox";
      bodyCc.placeholderText = "Type your notes here...";

      // Spacer
      body.insertParagraph("", Word.InsertLocation.end);
    });

    await context.sync();
  });
}

export { clearSapTemplate, generateSapFrontMatter, generateSapAfterToc };

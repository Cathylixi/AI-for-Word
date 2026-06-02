/* global Word */

import {
  TAG_PREFIX_MOCKTFL_TITLE,
  TAG_PREFIX_MOCKTFL_BODY,
  buildMockTflTitleTag,
  buildMockTflBodyTag
} from "./constants";
import { setupMockTflDocument } from "./mockTfl";
import { insertMockTflTitlePage } from "./titlePage";
import { insertConfidentialityStatement } from "./confidentiality";
import { insertMockTflSignaturePage } from "./signaturePage";
import { insertDocumentHistory } from "./documentHistory";
import { insertMockTflTableOfContents } from "./toc";
import { insertDividerPage } from "./divider";
import { setupMockTflPageHeader } from "./pageHeader";

// ─────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────

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

/**
 * Normalise and sort entries:
 *  - strip whitespace / cast types
 *  - sort: TABLE → FIGURE → LISTING, then by order field, then by number string
 */
function normalizeEntries(entries) {
  const typeOrder = { TABLE: 0, FIGURE: 1, LISTING: 2 };
  return (entries || [])
    .map((e) => ({
      order: typeof e?.order === "number" ? e.order : Number.MAX_SAFE_INTEGER,
      type: String(e?.type || "").trim().toUpperCase(),
      number: String(e?.number || "").trim(),
      title: String(e?.title || "").trim()
    }))
    .filter((e) => e.type && e.number && e.title)
    .sort((a, b) => {
      const ta = typeOrder[a.type] ?? 99;
      const tb = typeOrder[b.type] ?? 99;
      if (ta !== tb) return ta - tb;
      if (a.order !== b.order) return a.order - b.order;
      return a.number.localeCompare(b.number, "en");
    });
}

/**
 * Group entries by type (TABLE / FIGURE / LISTING), preserving typeOrder.
 */
function groupByType(entries) {
  const typeOrder = ["TABLE", "FIGURE", "LISTING"];
  const grouped = new Map(typeOrder.map((t) => [t, []]));
  entries.forEach((e) => {
    if (grouped.has(e.type)) grouped.get(e.type).push(e);
  });
  return { grouped, typeOrder };
}

// ─────────────────────────────────────────────────────────────
// Clear existing MockTFL content controls
// ─────────────────────────────────────────────────────────────

async function clearMockTflTemplate() {
  await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items");
    await context.sync();

    ccs.items.forEach((cc) => cc.load("tag"));
    await context.sync();

    const toDelete = ccs.items.filter((cc) => {
      const tag = cc.tag || "";
      return tag.startsWith(TAG_PREFIX_MOCKTFL_TITLE) || tag.startsWith(TAG_PREFIX_MOCKTFL_BODY);
    });

    for (let i = toDelete.length - 1; i >= 0; i--) {
      toDelete[i].delete(false);
    }
    await context.sync();
  });
}

// ─────────────────────────────────────────────────────────────
// Front matter  (Title → Conf → Signature → DocHistory → TOC)
// ─────────────────────────────────────────────────────────────

async function generateMockTflFrontMatter({ studyNumber }) {
  await Word.run(async (context) => {
    const body = context.document.body;

    // 1. Title Page
    insertMockTflTitlePage(body, {});
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

    // 2. Confidentiality Statement
    insertConfidentialityStatement(body);
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

    // 3. Signature Page
    insertMockTflSignaturePage(body);
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

    // 4. Document History (blank table)
    insertDocumentHistory(body);
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

    // 5. Table of Contents (lists selected entries with computed page numbers)
    insertMockTflTableOfContents(body);
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

    await context.sync();
  });
}

// ─────────────────────────────────────────────────────────────
// Full template generation
// ─────────────────────────────────────────────────────────────

async function generateMockTflTemplate({ studyNumber, entries }) {
  if (!studyNumber || !Array.isArray(entries) || entries.length === 0) return;

  const normalized = normalizeEntries(entries);
  if (normalized.length === 0) return;

  // Pre-group for divider pages.
  const { grouped, typeOrder } = groupByType(normalized);
  const nonEmptyTypes = typeOrder.filter((t) => (grouped.get(t) || []).length > 0);

  // 1. Clear document and set all sections to Landscape.
  await setupMockTflDocument();

  // 2. Front matter pages.
  await generateMockTflFrontMatter({ studyNumber });

  // 3. TFL body pages: DIVIDER → entries for each type.
  await Word.run(async (context) => {
    const body = context.document.body;

    for (let tIdx = 0; tIdx < nonEmptyTypes.length; tIdx++) {
      const type = nonEmptyTypes[tIdx];
      const typeItems = grouped.get(type) || [];
      const dividerText =
        type === "TABLE" ? "TABLES" : type === "FIGURE" ? "FIGURES" : "LISTINGS";

      // Divider page
      insertDividerPage(body, dividerText);
      body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

      typeItems.forEach((entry, idx) => {
        const number = String(entry.number || "").trim();
        const title = String(entry.title || "").trim();
        const line1 = `${entry.type} ${number}`.trim();

        // Title (single Content Control):
        // We want the display to match the user's screenshot:
        // - Line 1: "TABLE 14.1.1.2"
        // - Line 2: title text WITHOUT the last 3 words
        // - Line 3: the last 3 words (e.g., "Full Analysis Set")
        //
        // IMPORTANT:
        // We keep this as a single paragraph with soft line breaks (\u000B),
        // so Word shows only ONE folding triangle / one Heading 1 entry.
        const words = String(title || "")
          .trim()
          .split(/\s+/g)
          .filter(Boolean);

        // If the title has fewer than 4 words, we keep it as a single line (Line 2 only).
        let line2 = title;
        let line3 = "";
        if (words.length >= 4) {
          line2 = words.slice(0, -3).join(" ");
          line3 = words.slice(-3).join(" ");
        }

        const fullTitleText = line3
          ? `${line1}\u000B${line2}\u000B${line3}`
          : `${line1}\u000B${line2}`;
        const titlePara = body.insertParagraph(fullTitleText, Word.InsertLocation.end);
        
        // 1. Apply Style first (it sets defaults)
        applyHeadingStyle(titlePara, 1);

        // 2. Force formatting AFTER style to override defaults (e.g. blue color, left align)
        titlePara.alignment = "Centered";
        // IMPORTANT:
        // We keep Heading 1 semantics (navigation pane, outline) but override the visual style
        // to match the user's requested formatting (Courier New, size 8).
        titlePara.font.set({ name: "Courier New", bold: true, size: 8, color: "#000000" });
        try {
          titlePara.font.underline = "None";
        } catch (e) {}

        // Wrap paragraph in a content control
        const titleRange = titlePara.getRange();
        const titleCc = titleRange.insertContentControl();
        titleCc.tag = buildMockTflTitleTag({ studyNumber, type: entry.type, number });
        titleCc.title = `${line1} ${title}`.trim();
        titleCc.appearance = "BoundingBox";

        body.insertParagraph("", Word.InsertLocation.end);

        // Body placeholder content control
        const bodyPara = body.insertParagraph("", Word.InsertLocation.end);
        const bodyRange = bodyPara.getRange();
        bodyRange.font.set({ name: "Times New Roman", bold: false, size: 12 });
        const bodyCc = bodyRange.insertContentControl();
        bodyCc.tag = buildMockTflBodyTag({ studyNumber, type: entry.type, number });
        bodyCc.title = `${line1} ${title}`.trim();
        bodyCc.appearance = "BoundingBox";
        bodyCc.placeholderText = "Type your notes here...";

        // Page break between entries (and after last entry in a type group if more types follow)
        const isLastInType = idx === typeItems.length - 1;
        const isLastType = tIdx === nonEmptyTypes.length - 1;
        if (!isLastInType || !isLastType) {
          body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
        }
      });
    }

    // 4. Set up Page Headers (after all content is generated)
    await setupMockTflPageHeader(context);

    await context.sync();
  });
}

export { clearMockTflTemplate, generateMockTflFrontMatter, generateMockTflTemplate };

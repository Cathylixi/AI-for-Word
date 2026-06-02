/* global Word */

import { buildMockTflBodyTag, buildMockTflTitleTag } from "./constants";

/**
 * Read the Title Content Control text for a given MockTFL work item.
 *
 * Why we need this:
 * - The dialog should use the document as the source of truth.
 * - Users may edit the title in Word; reading from the title CC guarantees we analyze
 *   exactly what is in the document.
 */
async function readMockTflTitle({ studyNumber, type, number }) {
  return await Word.run(async (context) => {
    const tag = buildMockTflTitleTag({ studyNumber, type, number });
    const ccs = context.document.contentControls.getByTag(tag);
    ccs.load("items");
    await context.sync();
    if (ccs.items.length === 0) return "";
    const range = ccs.items[0].getRange();
    range.load("text");
    await context.sync();
    return range.text || "";
  });
}

/**
 * Read the third line of the MockTFL title (if available).
 * Used for dynamic footnotes (e.g. "Percentages are based on number of patients in [Line 3]").
 */
async function readMockTflTitleLine3({ studyNumber, type, number }) {
  const fullText = await readMockTflTitle({ studyNumber, type, number });
  const parts = String(fullText || "")
    .split(/\u000B|\r\n|\n/g)
    .map((s) => s.trim())
    .filter(Boolean);

  // In our 3-line template:
  // Line 0: TABLE 14.x.x
  // Line 1: Title Part 1
  // Line 2: Title Part 2 (e.g. "Full Analysis Set")
  if (parts.length >= 3) {
    return parts[2];
  }
  // Fallback: if only 2 lines, maybe line 1 contains the set name.
  if (parts.length === 2) {
    return parts[1];
  }
  return "";
}

/**
 * Extract the "pure title" that we want to embed for matching.
 *
 * In our template, the title CC contains:
 * - line 1: "TABLE 14.1.1.2"
 * - line 2: "Patient Disposition"            (example)
 * - line 3: "Full Analysis Set"             (example)
 *
 * We embed the semantic title lines (line 2+), so that:
 * - The "TABLE 14.1.1.2" prefix does not dominate the embedding
 * - Matching focuses on the semantic meaning of the title
 */
function extractPureTitleFromMockTflTitleCcText(titleCcText) {
  const raw = String(titleCcText || "");

  // Word may surface the soft line break (\u000B) or normal line breaks in different contexts.
  const parts = raw
    .split(/\u000B|\r\n|\n/g)
    .map((s) => String(s || "").trim())
    .filter(Boolean);

  let pure = "";
  if (parts.length >= 2) {
    // Template path: lines 2..end are the semantic title (we may split the title across multiple lines).
    pure = parts.slice(1).join(" ");
  } else {
    // Fallback path: try to remove "TABLE 14.x" prefix from a single-line title.
    pure = raw.replace(/^(TABLE|FIGURE|LISTING)\s+[0-9.]+\s+/i, "").trim();
  }

  // Normalize whitespace to avoid embedding variance due to formatting artifacts.
  pure = pure.replace(/\s+/g, " ").trim();
  return pure;
}

async function selectMockTflBody({ studyNumber, type, number }) {
  await Word.run(async (context) => {
    const tag = buildMockTflBodyTag({ studyNumber, type, number });
    const ccs = context.document.contentControls.getByTag(tag);
    ccs.load("items");
    await context.sync();
    if (ccs.items.length === 0) return;
    ccs.items[0].select();
    await context.sync();
  });
}

export {
  selectMockTflBody,
  readMockTflTitle,
  extractPureTitleFromMockTflTitleCcText,
  readMockTflTitleLine3
};

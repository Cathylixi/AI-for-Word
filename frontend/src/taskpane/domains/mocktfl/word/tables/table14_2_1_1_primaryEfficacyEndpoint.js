/* global Word */

import { buildMockTflBodyTag } from "../constants";
import { buildIndentedTableRowsFromEditorText } from "./parseMockTflEditorText";
import {
  buildColumnGroupHeaderText,
  buildMultiLevelHeaderRows,
  hasCustomSubLevels,
  resolveColumnHeaderConfig
} from "./columnHeaderConfig";
import { generateFixedLayoutTableOoxml } from "../ooxml/tableOoxml";
import { readMockTflTitleLine3 } from "../sections";

// Constants for the arm headers
const ARM_N_TEXT = "(N=xx)";

// Conservative fallback rows used only when the editor Notes are empty.
// 14.2.1.1 is a primary efficacy endpoint summary (time to first composite of
// VTE recurrence event), so the default skeleton focuses on the endpoint and
// its components.
const DEFAULT_ENDPOINT_ROWS = [
  "Number of patients with first composite VTE recurrence event",
  "DVT recurrence",
  "PE recurrence",
  "Other venous thromboembolic event",
  "Time to first composite VTE recurrence event"
];

/**
 * Build a minimal Notes-like text when the editor content is empty.
 * Keeping it as text lets us reuse the shared editor-text parser without
 * special-casing the row builder.
 */
function buildFallbackEditorText() {
  return DEFAULT_ENDPOINT_ROWS.join("\n");
}

/**
 * Build table rows for TABLE 14.2.1.1 (Primary Efficacy Endpoint / Full Analysis Set).
 *
 * Input expectation:
 * - `text` should be the final editable Notes content shown in the dialog.
 * - `selectedColumns` array of strings chosen by the user (e.g. ["Abelacimab", "Apixaban", "Overall"])
 */
function buildTableRows(text, selectedColumns = [], columnHeaderConfig = null) {
  const resolvedConfig = resolveColumnHeaderConfig({
    columnHeaderConfig,
    selectedColumns,
    defaultLevel1Headers: ["Abelacimab", "Apixaban", "Overall"]
  });
  const columnGroups = resolvedConfig.columnGroups;
  const columnGroupSubtitle = resolvedConfig.columnGroupSubtitle;

  // Only fall back to a default skeleton when the user has no Notes content.
  // We never overwrite AI-generated or user-edited Notes.
  const editorText = String(text || "").trim() ? String(text) : buildFallbackEditorText();

  if (hasCustomSubLevels(resolvedConfig)) {
    // Header layers expand beneath each selected column group.
    const { rows: headerRows, leafColumnCount } = buildMultiLevelHeaderRows({
      columnGroups,
      headerLayers: resolvedConfig.headerLayers,
      columnGroupSubtitle,
      armSuffixText: ARM_N_TEXT
    });
    return headerRows.concat(buildIndentedTableRowsFromEditorText(editorText, leafColumnCount + 1));
  }

  // Subtitle stays inside the same column-group cell, so it does not create an extra ruled header row.
  const headerRow1 = [
    "",
    ...columnGroups.map((c) =>
      buildColumnGroupHeaderText({
        columnGroup: c,
        armSuffixText: ARM_N_TEXT,
        columnGroupSubtitle
      })
    )
  ];
  const rows = [headerRow1];

  // columnCount is the question column (1) + selected arms
  return rows.concat(buildIndentedTableRowsFromEditorText(editorText, columnGroups.length + 1));
}

/**
 * Count leaf data columns from the generated header's first row.
 * Header cells may be objects carrying a colSpan, so we sum spans instead of length.
 */
function countLeafColumnsFromFirstRow(rows) {
  return Math.max(
    1,
    rows[0].slice(1).reduce((sum, cellValue) => {
      if (cellValue && typeof cellValue === "object" && !Array.isArray(cellValue)) {
        return sum + Math.max(1, Number(cellValue?.colSpan || 1));
      }
      return sum + 1;
    }, 0)
  );
}

/**
 * Build fixed column widths (twips) for the landscape mock shell.
 * Col 0 (statistic/label): 4.5 inch = 6480 twips.
 * Remaining 6480 twips is divided equally among the leaf arm columns.
 */
function buildColumnWidthsTwips(leafColumnCount) {
  const questionColWidth = 6480;
  const armColWidth = Math.floor(6480 / Math.max(1, leafColumnCount));
  return [questionColWidth, ...Array(leafColumnCount).fill(armColWidth)];
}

/**
 * Insert TABLE 14.2.1.1 table skeleton into the corresponding MockTFL body Content Control.
 * Includes footnotes:
 * - Percentages based on ...
 * - Source ...
 * - Program Name ...
 */
async function insertTable14_2_1_1_PrimaryEfficacyEndpoint({ studyNumber, text, selectedColumns, columnHeaderConfig }) {
  // 1. Get dynamic title part for footnote
  const populationText = await readMockTflTitleLine3({ studyNumber, type: "TABLE", number: "14.2.1.1" });

  await Word.run(async (context) => {
    const tag = buildMockTflBodyTag({
      studyNumber,
      type: "TABLE",
      number: "14.2.1.1"
    });

    const ccs = context.document.contentControls.getByTag(tag);
    ccs.load("items");
    await context.sync();
    if (ccs.items.length === 0) {
      throw new Error("Target body Content Control for TABLE 14.2.1.1 was not found.");
    }

    // Default to a 3-arm table if no columns were selected/provided
    const resolvedConfig = resolveColumnHeaderConfig({
      columnHeaderConfig,
      selectedColumns,
      defaultLevel1Headers: ["Abelacimab", "Apixaban", "Overall"]
    });
    const columnsToUse = resolvedConfig.columnGroups;
    const rows = buildTableRows(text, columnsToUse, resolvedConfig);
    // Header rows may span multiple leaf columns, so width allocation must use the expanded leaf count.
    const headerRowCount = hasCustomSubLevels(resolvedConfig) ? resolvedConfig.headerLayers.length + 1 : 1;
    const placeholderCc = ccs.items[0];

    const leafColumnCount = countLeafColumnsFromFirstRow(rows);
    const columnWidthsTwips = buildColumnWidthsTwips(leafColumnCount);

    const ooxml = generateFixedLayoutTableOoxml({
      rows,
      columnWidthsTwips,
      useCourier8: true,
      headerRowCount
    });

    // Strategy (Simplified & Direct):
    // 1. Clear the placeholder CC content.
    // 2. Insert OOXML Table INSIDE the CC (at Start).
    // 3. Insert Footnotes INSIDE the CC (after the table).

    placeholderCc.clear();
    await context.sync();

    // Insert at Start of the now-empty CC.
    const insertedRange = placeholderCc.insertOoxml(ooxml, Word.InsertLocation.start);
    await context.sync();

    // Insert footnotes using the range after the table.
    const afterTableRange = insertedRange.getRange("After");

    const fp1 = afterTableRange.insertParagraph(
      `Percentages are based on number of patients in ${populationText || "Full Analysis Set"}.`,
      Word.InsertLocation.after
    );
    fp1.font.set({ name: "Courier New", size: 8, bold: false });

    const fp2 = fp1.insertParagraph("Source: Listing 16.2.X.X", Word.InsertLocation.after);
    fp2.font.set({ name: "Courier New", size: 8, bold: false });

    // Footnote 3: Program Name / DB / Runtime (Paragraph with Tab Stops)
    // We use tabs to align: Left (Program), Center (DB), Right (Runtime).
    // Assumes Landscape page width approx 9 inches (648 pts).
    const footerText = "Program Name: xxxxxxxxxxxx.sas\tDB <Snapshot/Lock> Date: DDMMMYYYY\tRuntime: DDMMMYYYY HH:MM";
    const p3 = fp2.insertParagraph(footerText, Word.InsertLocation.after);
    p3.font.set({ name: "Courier New", size: 8, bold: false });

    // Set Tab Stops
    // Center tab at 4.5 inch (324 pt)
    // Right tab at 9.0 inch (648 pt)
    try {
      const tabs = p3.tabStops;
      tabs.clearAll();
      tabs.add({ position: 324, alignment: Word.TabAlignment.center, leader: Word.TabLeader.none });
      tabs.add({ position: 648, alignment: Word.TabAlignment.right, leader: Word.TabLeader.none });
    } catch (e) {}

    await context.sync();
  });
}

export { insertTable14_2_1_1_PrimaryEfficacyEndpoint };

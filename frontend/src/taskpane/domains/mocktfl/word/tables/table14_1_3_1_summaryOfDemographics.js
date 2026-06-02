/* global Word */

import { buildMockTflBodyTag } from "../constants";
import { buildIndentedTableRowsFromEditorText } from "./parseMockTflEditorText";
import {
  buildColumnGroupHeaderText,
  buildMultiLevelHeaderRows,
  hasCustomSubLevels,
  resolveColumnHeaderConfig
} from "./columnHeaderConfig";

// Constants for the arm headers
const ARM_N_TEXT = "(N=xx)";

/**
 * Build table rows for TABLE 14.1.3.1 (Summary of demographics / Full Analysis Set).
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

  if (hasCustomSubLevels(resolvedConfig)) {
    // Header layers expand beneath each selected column group.
    const { rows: headerRows, leafColumnCount } = buildMultiLevelHeaderRows({
      columnGroups,
      headerLayers: resolvedConfig.headerLayers,
      columnGroupSubtitle,
      armSuffixText: ARM_N_TEXT
    });
    return headerRows.concat(buildIndentedTableRowsFromEditorText(text, leafColumnCount + 1));
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
  return rows.concat(buildIndentedTableRowsFromEditorText(text, columnGroups.length + 1));
}

/**
 * Apply the same visual formatting rules as TABLE 14.1.1.2:
 * - 50/50 split width (question column vs arm columns)
 * - Courier New, size 8
 * - Question column left aligned; arm/value columns centered
 * - Only 3 horizontal lines (top, under header, bottom)
 *
 * IMPORTANT:
 * - Many Word APIs only throw when calling `context.sync()`.
 * - We apply formatting in small steps and catch sync failures per step.
 */
async function formatDemographicsTable(context, table, rows) {
  // 0) Set explicit column widths (4.5 inches for col 0, 1.5 inches for cols 1-3).
  // IMPORTANT:
  // - In many Word hosts, `TableColumn.width` is the most reliable writable property.
  // - `preferredWidth` may be ignored depending on Word/Office build and table style.
  // - `TableColumn.setWidth()` is an official API that sometimes applies more reliably than property assignment.
  // So we use setWidth() when available, and also set width/preferredWidth as best-effort fallbacks.
  //
  // 1 inch = 72 points.
  // Col 0: 4.5 * 72 = 324 points
  // Col 1,2,3: 1.5 * 72 = 108 points
  try {
    // Disable autofit so we can set fixed widths.
    try {
      if (typeof table.autoFitBehavior === "function") {
        const fixed =
          (Word && Word.TableAutoFitBehavior && Word.TableAutoFitBehavior.fixedColumnWidth) || "FixedColumnWidth";
        table.autoFitBehavior(fixed);
      }
    } catch (e) {}

    const c0 = table.columns.getItemAt(0);
    const c1 = table.columns.getItemAt(1);
    const c2 = table.columns.getItemAt(2);
    const c3 = table.columns.getItemAt(3);

    // Preferred path (official API): setWidth in Points.
    try {
      const pointsStyle = (Word && Word.RulerStyle && Word.RulerStyle.points) || "Points";
      if (typeof c0.setWidth === "function") c0.setWidth(324, pointsStyle);
      if (typeof c1.setWidth === "function") c1.setWidth(108, pointsStyle);
      if (typeof c2.setWidth === "function") c2.setWidth(108, pointsStyle);
      if (typeof c3.setWidth === "function") c3.setWidth(108, pointsStyle);
    } catch (e) {}

    // Fallback path: set explicit column widths by property assignment.
    try {
      c0.width = 324;
      c1.width = 108;
      c2.width = 108;
      c3.width = 108;
    } catch (e) {}

    // Secondary (best-effort) path: also set preferred widths when supported.
    try {
      c0.preferredWidth = 324;
      c1.preferredWidth = 108;
      c2.preferredWidth = 108;
      c3.preferredWidth = 108;
    } catch (e) {}

    await context.sync();
  } catch (e) {
    // If width APIs fail, we keep default Word layout behavior.
  }

  // 1) Remove all borders first (table-level), then cell-level (style-driven borders).
  try {
    ["Top", "Bottom", "Left", "Right", "InsideHorizontal", "InsideVertical"].forEach((b) => {
      table.getBorder(b).type = "None";
    });
    await context.sync();
  } catch (e) {}

  try {
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < 4; c++) {
        const cell = table.getCell(r, c);
        ["Top", "Bottom", "Left", "Right"].forEach((b) => {
          try {
            cell.getBorder(b).type = "None";
          } catch (e) {}
        });
      }
    }
    await context.sync();
  } catch (e) {}

  // 2) Font + alignment.
  try {
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < 4; c++) {
        const cellRange = table.getCell(r, c).body.getRange();
        cellRange.font.set({ name: "Courier New", size: 8, bold: false });
        try {
          cellRange.paragraphFormat.alignment = c === 0 ? "Left" : "Centered";
        } catch (e) {}
      }
    }
    await context.sync();
  } catch (e) {}

  // 3) Draw only the 3 required horizontal lines.
  try {
    for (let c = 0; c < 4; c++) {
      table.getCell(0, c).getBorder("Top").type = "Single";
      table.getCell(1, c).getBorder("Bottom").type = "Single";
    }
    const lastRow = Math.max(0, rows.length - 1);
    for (let c = 0; c < 4; c++) {
      table.getCell(lastRow, c).getBorder("Bottom").type = "Single";
    }
    await context.sync();
  } catch (e) {}
}

import { generateFixedLayoutTableOoxml } from "../ooxml/tableOoxml";
import { readMockTflTitleLine3 } from "../sections";

/**
 * Insert TABLE 14.1.3.1 table skeleton into the corresponding MockTFL body Content Control.
 * Includes footnotes:
 * - Percentages based on ...
 * - Source ...
 * - Program Name ...
 */
async function insertTable14_1_3_1_SummaryOfDemographics({ studyNumber, text, selectedColumns, columnHeaderConfig }) {
  // 1. Get dynamic title part for footnote
  const populationText = await readMockTflTitleLine3({ studyNumber, type: "TABLE", number: "14.1.3.1" });

  await Word.run(async (context) => {
    const tag = buildMockTflBodyTag({
      studyNumber,
      type: "TABLE",
      number: "14.1.3.1"
    });

    const ccs = context.document.contentControls.getByTag(tag);
    ccs.load("items");
    await context.sync();
    if (ccs.items.length === 0) {
      throw new Error("Target body Content Control for TABLE 14.1.3.1 was not found.");
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

    // IMPORTANT (OOXML implementation):
    // Widths in twips (1 inch = 1440 twips):
    // Col 0 (Questions): 4.5 inch = 6480 twips
    // Remaining 6480 twips is divided equally among the selected arms.
    const questionColWidth = 6480;
    const leafColumnCount = Math.max(
      1,
      rows[0].slice(1).reduce((sum, cellValue) => {
        if (cellValue && typeof cellValue === "object" && !Array.isArray(cellValue)) {
          return sum + Math.max(1, Number(cellValue?.colSpan || 1));
        }
        return sum + 1;
      }, 0)
    );
    const armColWidth = Math.floor(6480 / leafColumnCount);
    const columnWidthsTwips = [questionColWidth, ...Array(leafColumnCount).fill(armColWidth)];

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

    const fp2 = fp1.insertParagraph("Source: Listing 16.2.1.X", Word.InsertLocation.after);
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
    } catch(e) {}

    await context.sync();
  });
}

export { insertTable14_1_3_1_SummaryOfDemographics };

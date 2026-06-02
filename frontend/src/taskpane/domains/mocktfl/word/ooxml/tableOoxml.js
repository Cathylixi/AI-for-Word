/**
 * Generate a Word OOXML string for a table with fixed column widths.
 *
 * Why OOXML?
 * Office.js APIs for column width (preferredWidth, width, setWidth) are often overridden
 * by Word's auto-fit behavior or document defaults. OOXML allows us to define
 * <w:tblLayout w:type="fixed"/> and explicit <w:gridCol> widths (in twips),
 * which Word treats as definitive.
 *
 * Units:
 * 1 inch = 1440 twips.
 */

// Escape XML special characters
function escapeXml(unsafe) {
  return String(unsafe || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;")
    // Soft line break (\u000B) -> valid OOXML run break <w:br/>
    // Note: In Word OOXML, breaks are elements, not text chars. We handle this in cell generation.
    .replace(/\u000B/g, "\u000B");
}

/**
 * @param {Object} params
 * @param {string[][]} params.rows - 2D array of cell text.
 * @param {number[]} params.columnWidthsTwips - Array of widths in twips (e.g. [6480, 2160...]).
 * @param {boolean} [params.useCourier8] - If true, applies Courier New 8pt to all cells.
 * @param {number} [params.headerRowCount=2] - Number of leading header rows.
 */
export function generateFixedLayoutTableOoxml({ rows, columnWidthsTwips, useCourier8 = true, headerRowCount = 2 }) {
  const totalWidth = columnWidthsTwips.reduce((a, b) => a + b, 0);
  const safeHeaderRowCount = Math.max(1, Number(headerRowCount || 1));

  // XML Header
  let xml = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:tbl>
            <w:tblPr>
              <w:tblStyle w:val="TableGrid"/>
              <w:tblW w:w="0" w:type="auto"/>
              <w:tblLayout w:type="fixed"/>
              <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
            </w:tblPr>
            <w:tblGrid>`;

  // Define Grid Columns (The source of truth for column widths)
  columnWidthsTwips.forEach((w) => {
    xml += `<w:gridCol w:w="${w}"/>`;
  });
  xml += `</w:tblGrid>`;

  // Generate Rows
  rows.forEach((row, rIndex) => {
    xml += `<w:tr>`;
    // Row height (optional, auto is usually fine, but we can set it if needed)
    // <w:trPr><w:trHeight w:val="288"/></w:trPr>

    let logicalColumnIdx = 0;
    row.forEach((cellValue) => {
      const cell =
        cellValue && typeof cellValue === "object" && !Array.isArray(cellValue)
          ? cellValue
          : { text: cellValue, colSpan: 1 };
      const colSpan = Math.max(1, Number(cell?.colSpan || 1));
      const cIndex = logicalColumnIdx;
      const width = columnWidthsTwips.slice(logicalColumnIdx, logicalColumnIdx + colSpan).reduce((sum, next) => sum + next, 0) || 2160;
      
      // Determine if this cell needs special OOXML indenting.
      // We look for the "__INDENT__" prefix injected by our parser.
      const isIndented = typeof cell?.text === "string" && cell.text.startsWith("__INDENT__");
      const cellTextToRender = isIndented ? cell.text.replace("__INDENT__", "") : cell?.text;

      // Determine borders:
      // We replicate the "3 horizontal lines" design in OOXML to ensure it sticks.
      // - Top of Row 0: Single
      // - Bottom of Row 1: Single
      // - Bottom of Last Row: Single
      // - All others: None
      // Note: "TableGrid" style usually adds borders. We explicitly override cell borders here.
      
      let topVal = "none";
      let bottomVal = "none";
      let leftVal = "none";
      let rightVal = "none";

      if (rIndex === 0) topVal = "single";
      // 横向分隔多层表头 (只在有内容的列下方加线，保持左上角空白处无底线)
      if (rIndex < safeHeaderRowCount - 1 && cIndex > 0) bottomVal = "single";
      if (rIndex === safeHeaderRowCount - 1) bottomVal = "single";
      if (rIndex === rows.length - 1) bottomVal = "single";

      // Font style XML fragment
      // Courier New, size 8 (16 half-points)
      // Alignment: Column 0 = Left, Others = Center
      const align = cIndex === 0 ? "left" : "center";
      const runProps = useCourier8 
        ? `<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>`
        : "";

      xml += `<w:tc>
        <w:tcPr>
          <w:tcW w:w="${width}" w:type="dxa"/>
          <!-- gridSpan lets a level-1/2 header visually span its repeated child columns -->
          ${colSpan > 1 ? `<w:gridSpan w:val="${colSpan}"/>` : ""}
          <w:tcBorders>
            <w:top w:val="${topVal}" w:sz="4" w:space="0" w:color="auto"/>
            <w:left w:val="${leftVal}" w:sz="4" w:space="0" w:color="auto"/>
            <w:bottom w:val="${bottomVal}" w:sz="4" w:space="0" w:color="auto"/>
            <w:right w:val="${rightVal}" w:sz="4" w:space="0" w:color="auto"/>
          </w:tcBorders>
          <w:vAlign w:val="center"/>
        </w:tcPr>
        <w:p>
          <w:pPr>
            <w:jc w:val="${align}"/>
            ${isIndented ? '<w:ind w:left="288" />' : ''}
            <!-- Line spacing:
                 w:line="240" is single spacing (240 twips = 12pt).
                 We set w:before="0" and w:after="0" to remove Word's default paragraph gaps
                 so the table looks compact. Visual spacing between blocks is achieved
                 by explicit blank rows instead of paragraph margins.
            -->
            <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
          </w:pPr>`;

      // Handle text with possible soft breaks (\u000B)
      const parts = escapeXml(cellTextToRender).split("\u000B");
      parts.forEach((part, i) => {
        if (i > 0) {
          xml += `<w:r>${runProps}<w:br/></w:r>`;
        }
        xml += `<w:r>${runProps}<w:t>${part}</w:t></w:r>`;
      });

      xml += `</w:p></w:tc>`;
      logicalColumnIdx += colSpan;
    });
    xml += `</w:tr>`;
  });

  xml += `</w:tbl>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

  return xml;
}

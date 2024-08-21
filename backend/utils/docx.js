const { Document, Packer } = require("docx");
const fs = require("fs");

async function extractStylesFromDocx(filePath) {
  // Load the DOCX file
  const doc = await new Document(fs.readFileSync(filePath));

  // Initialize an array to store the extracted data
  const extractedData = [];

  // Traverse through the document's children
  doc.body.children.forEach((child) => {
    if (child instanceof Paragraph) {
      const paragraphData = extractParagraphStyles(child);
      extractedData.push(paragraphData);
    } else if (child instanceof Table) {
      const tableData = extractTableStyles(child);
      extractedData.push(tableData);
    }
    // Handle other types (e.g., images) if needed
  });

  return extractedData;
}

function extractParagraphStyles(paragraph) {
  const runs = paragraph.children.map((run) => {
    return {
      text: run.text,
      style: run.style, // This includes color, bold, italic, etc.
    };
  });

  return {
    type: "paragraph",
    alignment: paragraph.alignment,
    margin: paragraph.margin, // Extract margins if available
    runs: runs,
  };
}

function extractTableStyles(table) {
  const rows = table.rows.map((row) => {
    return row.cells.map((cell) => {
      return {
        text: cell.text,
        style: cell.style, // Cell styles, like border and padding
      };
    });
  });

  return {
    type: "table",
    rows: rows,
  };
}

module.exports = { extractStylesFromDocx };

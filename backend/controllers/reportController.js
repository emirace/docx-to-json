const {
  processDocxFile,
  extractDocxContent,
} = require("../utils/fileProcessor");
const { extractStylesFromDocx } = require("../utils/docx");
const Report = require("../models/report");
const path = require("path");
const fs = require("fs");

const uploadReport = async (req, res) => {
  const filePath = req.file.path;
  // const jsonData = await processDocxFile(filePath);
  // const jsonData = await extractStylesFromDocx(filePath);
  // Example usage
  const jsonData = await extractDocxContent(filePath);

  // const report = new Report({ jsonData });
  // await report.save();

  fs.unlinkSync(filePath); // Clean up the uploaded file

  res.json(jsonData);
};

module.exports = { uploadReport };

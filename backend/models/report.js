const mongoose = require("mongoose");

const ReportSchema = new mongoose.Schema({
  jsonData: { type: Object, required: true },
  images: [String],
});

module.exports = mongoose.model("Report", ReportSchema);

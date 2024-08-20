const mammoth = require("mammoth");
const cheerio = require("cheerio");
const fs = require("fs");

const processDocxFile = async (filePath) => {
  const buffer = fs.readFileSync(filePath);

  // Custom style map
  const styleMap = [
    "p[style-name='Cover Product title'] => p.cover-product-title",
    "p[style-name='Document title'] => p.document-title",
    "p[style-name='Cover Job details'] => p.cover-job-details",
    "p[style-name='TOC Heading'] => p.toc-heading",
    "p[style-name='TGT_Body'] => p.tgt-body",
    "p[style-name='toc 1'] => p.toc-1",
    "p[style-name='toc 2'] => p.toc-2",
    "p[style-name='TGT_HEADING2'] => h2.tgt-heading2",
    "p[style-name='List Bullet'] => ul > li",
    "p[style-name='caption'] => p.caption",
    "p[style-name='TGT_Caption'] => p.tgt-caption",
    "p[style-name='TGT_HEADING1'] => h1.tgt-heading1",
    "p[style-name='Body Text'] => p.body-text",
    "r[style-name='Body Text Char'] => span.body-text-char",
    "r[style-name='Caption Char'] => span.caption-char",
    "r[style-name='Heading 1 Char'] => span.heading1-char",
  ];

  const result = await mammoth.convertToHtml({ buffer }, { styleMap });

  // Log any warnings
  result.messages.forEach((message) => {
    console.warn(message);
  });

  const jsonData = parseHtmlToJson(result.value);
  return jsonData;
};

const parseHtmlToJson = (htmlContent) => {
  const $ = cheerio.load(htmlContent);

  const buildJsonStructure = (element) => {
    const children = [];

    $(element)
      .children()
      .each((_, child) => {
        children.push(buildJsonStructure(child));
      });

    return {
      tag: $(element).get(0).tagName,
      text: $(element).text(),
      style: $(element).attr("class") || "",
      children: children.length > 0 ? children : undefined,
    };
  };
  const jsonResult = buildJsonStructure($.root());

  return jsonResult;
};

module.exports = { processDocxFile };

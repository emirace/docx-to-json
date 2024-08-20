const mammoth = require("mammoth");
const cheerio = require("cheerio");

const processDocxFile = async (filePath) => {
  const { value: htmlContent, messages } = await mammoth.convertToHtml({
    path: filePath,
  });

  //   console.log(htmlContent);
  console.log(messages);
  const jsonData = parseHtmlToJson(htmlContent);
  return jsonData;
};

const parseHtmlToJson = (htmlContent) => {
  const $ = cheerio.load(htmlContent);

  const headers = [];
  const paragraphs = [];
  const lists = [];
  const images = [];
  const tables = [];

  $("h1, h2, h3, h4, h5, h6").each((i, el) => {
    console.log($(el).attr());
    headers.push({
      tag: el.tagName,
      text: $(el).text(),
      color: $(el).css("color"),
      style: $(el).attr("style"),
    });
  });

  $("p").each((i, el) => {
    paragraphs.push({
      text: $(el).text(),
      style: $(el).attr("style"),
      color: $(el).css("color"),
    });
  });

  $("ul, ol").each((i, el) => {
    const items = [];
    $(el)
      .find("li")
      .each((j, li) => {
        items.push($(li).text());
      });
    lists.push({
      type: el.tagName,
      items,
    });
  });

  $("img").each((i, el) => {
    images.push({
      src: $(el).attr("src"),
      alt: $(el).attr("alt"),
      caption: $(el).next("figcaption").text(),
    });
  });

  $("table").each((i, el) => {
    const rows = [];
    $(el)
      .find("tr")
      .each((j, row) => {
        const cells = [];
        $(row)
          .find("td, th")
          .each((k, cell) => {
            cells.push($(cell).text());
          });
        rows.push(cells);
      });
    tables.push({
      caption: $(el).prev("caption").text(),
      rows,
    });
  });

  return {
    headers,
    paragraphs,
    lists,
    images,
    tables,
  };
};

module.exports = { processDocxFile };

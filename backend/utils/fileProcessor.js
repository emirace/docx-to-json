const mammoth = require("mammoth");
const cheerio = require("cheerio");
const path = require("path");
const unzipper = require("unzipper");
const fs = require("fs");
const rimraf = require("rimraf");

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

  const result = await mammoth.extractRawText({ buffer }, { styleMap });

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

async function extractDocxContent(filePath) {
  try {
    const outputDir = path.join(__dirname, "output");
    await fs.promises.mkdir(outputDir, { recursive: true });

    await fs
      .createReadStream(filePath)
      .pipe(unzipper.Extract({ path: outputDir }))
      .promise();

    const documentXmlPath = path.join(outputDir, "word", "document.xml");
    const stylesXmlPath = path.join(outputDir, "word", "styles.xml");
    const numberingXmlPath = path.join(outputDir, "word", "numbering.xml");

    const documentXml = await fs.promises.readFile(documentXmlPath, "utf8");
    const stylesXml = await fs.promises.readFile(stylesXmlPath, "utf8");
    const numberingXml = await fs.promises.readFile(numberingXmlPath, "utf8");

    const $doc = cheerio.load(documentXml, { xmlMode: true });
    const $styles = cheerio.load(stylesXml, { xmlMode: true });
    const $numbering = cheerio.load(numberingXml, { xmlMode: true });

    // Build numbering map from numbering.xml
    const numberingMap = buildNumberingMap($numbering);

    // console.log(JSON.stringify(numberingMap));

    const styleMap = {};
    $styles("w\\:style").each((_, style) => {
      const styleId = $styles(style).attr("w:styleId");
      const styleType = $styles(style).attr("w:type");

      if (styleId && styleType) {
        styleMap[styleId] = {
          type: styleType,
          name: $styles(style).find("w\\:name").attr("w:val"),
          basedOn: $styles(style).find("w\\:basedOn").attr("w:val"), // Capture the basedOn attribute
          runProperties: extractRunStyles($styles(style).find("w\\:rPr")),
          paragraphProperties: extractParagraphStyles(
            $styles(style).find("w\\:pPr")
          ),
        };
      }
    });

    // Resolve inherited styles based on the 'basedOn' attribute
    Object.keys(styleMap).forEach((styleId) => {
      const style = styleMap[styleId];
      if (style.basedOn) {
        const inheritedStyle = styleMap[style.basedOn];
        if (inheritedStyle) {
          style.runProperties = {
            ...inheritedStyle.runProperties,
            ...style.runProperties,
          };
          style.paragraphProperties = {
            ...inheritedStyle.paragraphProperties,
            ...style.paragraphProperties,
          };
        }
      }
    });

    function extractRunStyles(rPr) {
      const styles = {};

      if (rPr.find("w\\:b").length > 0) styles.bold = true;
      if (rPr.find("w\\:i").length > 0) styles.italic = true;
      if (rPr.find("w\\:u").length > 0) styles.underline = true;
      if (rPr.find("w\\:strike").length > 0) styles.strikeThrough = true;

      const color = rPr.find("w\\:color").attr("w:val");
      if (color) styles.color = color;

      const fontSize = rPr.find("w\\:sz").attr("w:val");
      if (fontSize) styles.fontSize = fontSize;

      const font = rPr.find("w\\:rFonts").attr("w:ascii");
      if (font) styles.font = font;

      const backgroundColor = rPr.find("w\\:shd").attr("w:fill");
      if (backgroundColor) styles.backgroundColor = backgroundColor;

      const highlight = rPr.find("w\\:highlight").attr("w:val");
      if (highlight) styles.highlight = highlight;

      return styles;
    }

    function extractParagraphStyles(pPr) {
      const styles = {};

      const alignment = pPr.find("w\\:jc").attr("w:val");
      if (alignment) styles.alignment = alignment;

      const spacingBefore = pPr.find("w\\:spacing").attr("w:before");
      if (spacingBefore) styles.spacingBefore = spacingBefore;

      const spacingAfter = pPr.find("w\\:spacing").attr("w:after");
      if (spacingAfter) styles.spacingAfter = spacingAfter;

      const indentLeft = pPr.find("w\\:ind").attr("w:left");
      if (indentLeft) styles.indentLeft = indentLeft;

      const indentRight = pPr.find("w\\:ind").attr("w:right");
      if (indentRight) styles.indentRight = indentRight;

      return styles;
    }

    function parseElement(element) {
      const children = [];

      element.children().each((_, child) => {
        const tag = $doc(child)[0].tagName;

        if (tag === "w:p") {
          const paragraphData = {
            type: "paragraph",
            text: "",
            styles: {},
          };

          const pPr = $doc(child).find("w\\:pPr");
          const pStyleId = pPr.find("w\\:pStyle").attr("w:val");

          if (pStyleId && styleMap[pStyleId]) {
            paragraphData.styles = {
              ...styleMap[pStyleId].paragraphProperties,
              ...styleMap[pStyleId].runProperties,
            };
          }

          if (pPr.length) {
            paragraphData.styles = {
              ...paragraphData.styles,
              ...extractParagraphStyles(pPr),
            };

            const numPr = pPr.find("w\\:numPr");
            if (numPr.length > 0) {
              const numId = numPr.find("w\\:numId").attr("w:val");
              const ilvl = numPr.find("w\\:ilvl").attr("w:val");

              const listData = extractListInfo(numId, ilvl, numberingMap);
              paragraphData.listData = listData;
            }
          }

          const nextChild = [];
          $doc(child)
            .find("w\\:r, w\\:drawing, w\\:pict")
            .each((_, run) => {
              const runTag = $doc(run)[0].tagName;

              if (runTag === "w:r") {
                const runText = $doc(run).find("w\\:t").text();
                const rPr = $doc(run).find("w\\:rPr");
                const rStyleId = rPr.find("w\\:rStyle").attr("w:val");
                let runStyles = {};

                if (rStyleId && styleMap[rStyleId]) {
                  runStyles = {
                    ...styleMap[rStyleId].runProperties,
                  };
                }
                runStyles = {
                  ...runStyles,
                  ...extractRunStyles(rPr),
                };
                nextChild.push({
                  text: runText,
                  styles: runStyles,
                });
              } else if (runTag === "w:drawing") {
                const imageData = parseDrawing(run);
                children.push(imageData);
              }
            });

          paragraphData.text = nextChild
            .filter((child) => child.text)
            .map((child) => child.text)
            .join("");
          paragraphData.styles = {
            ...paragraphData.styles,
            ...nextChild.styles,
          };
          paragraphData.styleName = pStyleId;
          children.push(paragraphData);
        } else if (tag === "w:tbl") {
          const tableData = {
            type: "table",
            rows: [],
          };

          $doc(child)
            .find("w\\:tr")
            .each((_, row) => {
              const rowData = [];

              $doc(row)
                .find("w\\:tc")
                .each((_, cell) => {
                  const cellData = {
                    type: "cell",
                    content: parseElement($doc(cell)),
                  };
                  rowData.push(cellData);
                });

              tableData.rows.push(rowData);
            });

          children.push(tableData);
        } else if (tag === "w:sectPr") {
          const sectionData = {
            type: "section",
            styles: {
              pageSize:
                $doc(child).find("w\\:pgSz").attr("w:w") +
                "x" +
                $doc(child).find("w\\:pgSz").attr("w:h"),
              margins: {
                top: $doc(child).find("w\\:pgMar").attr("w:top"),
                bottom: $doc(child).find("w\\:pgMar").attr("w:bottom"),
                left: $doc(child).find("w\\:pgMar").attr("w:left"),
                right: $doc(child).find("w\\:pgMar").attr("w:right"),
              },
            },
          };
          children.push(sectionData);
        }
      });

      return children;
    }

    function buildNumberingMap($numbering) {
      const abstractNumMap = {}; // Stores the abstract numbering definitions
      const numberingMap = {}; // Stores the final mapping of numId to its levels

      // Recursive function to resolve abstract numbering, including any references to other abstractNumIds
      function resolveAbstractNum(abstractNumId) {
        // If this abstractNumId is already resolved, return it
        if (abstractNumMap[abstractNumId]?.resolved) {
          return abstractNumMap[abstractNumId].levels;
        }

        const levels = abstractNumMap[abstractNumId]?.levels || [];

        // Check if this abstractNumId refers to another abstractNumId
        const referencedAbstractNumId =
          abstractNumMap[abstractNumId]?.referencedAbstractNumId;
        if (referencedAbstractNumId) {
          // Recursively resolve the referenced abstractNumId and merge the levels
          const referencedLevels = resolveAbstractNum(referencedAbstractNumId);
          Object.assign(levels, referencedLevels);
        }

        // Mark this abstractNumId as resolved to avoid circular references
        abstractNumMap[abstractNumId] = {
          levels,
          resolved: true,
        };

        return levels;
      }

      // Loop through each abstract numbering definition and build the abstractNumMap
      $numbering("w\\:abstractNum").each((_, abstractNum) => {
        const abstractNumId = $numbering(abstractNum).attr("w:abstractNumId");
        const levels = [];

        $numbering(abstractNum)
          .find("w\\:lvl")
          .each((_, level) => {
            const levelIndex = $numbering(level).attr("w:ilvl");
            const numFmt = $numbering(level).find("w\\:numFmt").attr("w:val");
            const lvlText = $numbering(level).find("w\\:lvlText").attr("w:val");

            // Get indentation spacing
            const indent = {};
            const indElement = $numbering(level).find("w\\:ind");
            if (indElement.length > 0) {
              indent.left = indElement.attr("w:left") || null;
              indent.hanging = indElement.attr("w:hanging") || null;
              indent.firstLine = indElement.attr("w:firstLine") || null;
            }

            levels[levelIndex] = {
              numFmt,
              lvlText,
              indent,
            };
          });

        // Check if this abstractNum refers to another abstractNum
        const referencedAbstractNumId = $numbering(abstractNum)
          .find("w\\:nsid")
          .attr("w:val");
        abstractNumMap[abstractNumId] = {
          levels,
          referencedAbstractNumId: referencedAbstractNumId || null,
          resolved: false, // Initially set as unresolved
        };
      });

      // Loop through each num element to build the numberingMap
      $numbering("w\\:num").each((_, num) => {
        const numId = $numbering(num).attr("w:numId");
        const abstractNumId = $numbering(num)
          .find("w\\:abstractNumId")
          .attr("w:val");

        if (abstractNumId) {
          numberingMap[numId] = resolveAbstractNum(abstractNumId);
        }
      });

      return numberingMap;
    }

    function parseDrawing(drawingElement) {
      const anchor = $doc(drawingElement).find("wp\\:anchor");
      const graphicData = $doc(drawingElement).find("a\\:graphicData");

      const imageData = {
        type: "image",
        src: null,
        position: {
          horizontal: null,
          vertical: null,
        },
        size: {
          width: null,
          height: null,
        },
        properties: {},
      };

      const positionH = anchor.find("wp\\:positionH wp\\:align").text();
      const positionV = anchor.find("wp\\:positionV wp\\:align").text();
      const cx = anchor.find("wp\\:extent").attr("cx");
      const cy = anchor.find("wp\\:extent").attr("cy");

      if (positionH && positionV) {
        imageData.position.horizontal = positionH;
        imageData.position.vertical = positionV;
      }

      if (cx && cy) {
        imageData.size.width = parseInt(cx, 10);
        imageData.size.height = parseInt(cy, 10);
      }

      const blip = graphicData.find("a\\:blip");
      const embedId = blip.attr("r:embed");

      if (embedId) {
        const relsPath = path.join(
          outputDir,
          "word",
          "_rels",
          "document.xml.rels"
        );
        const relsXml = fs.readFileSync(relsPath, "utf8");
        const $rels = cheerio.load(relsXml, { xmlMode: true });

        const target = $rels(`Relationship[Id="${embedId}"]`).attr("Target");
        if (target) {
          const imagePath = path.join(outputDir, "word", target);
          const imageBuffer = fs.readFileSync(imagePath);
          const base64Image = `data:image/png;base64,${imageBuffer.toString(
            "base64"
          )}`;
          imageData.src = base64Image;
        }
      }

      return imageData;
    }

    function extractListInfo(numId, ilvl, numberingMap) {
      const levelInfo = numberingMap[numId] && numberingMap[numId][ilvl];

      return levelInfo
        ? {
            isBullet: levelInfo.numFmt === "bullet",
            bulletText: levelInfo.lvlText,
            level: ilvl,
            indent: levelInfo.indent,
          }
        : null;
    }

    const parsedContent = parseElement($doc("w\\:document > w\\:body"));
    rimraf.sync(outputDir);

    return mapSections(parsedContent);
  } catch (err) {
    console.error("Error extracting DOCX content:", err);
    throw err;
  }
}

function mapSections(paragraphs) {
  const sections = [];
  let currentSection = null;
  let currentSubsection = null;
  let preamble = { body: [] };

  paragraphs.forEach((paragraph) => {
    const { styleName, text } = paragraph;

    if (["Heading1", "TGTHEADING1"].includes(styleName)) {
      if (currentSubsection) {
        currentSection.body.push(currentSubsection);
      }

      // If there's an active section, push it to the sections array
      if (currentSection) {
        sections.push(currentSection);
      }

      // Start a new main section
      currentSection = {
        title: text,
        body: [],
        // subsections: [],
      };

      // Reset current subsection
      currentSubsection = null;
    } else if (styleName === "TGTHEADING2" && currentSection) {
      // If there's an active subsection, push it to the subsections array
      if (currentSubsection) {
        currentSection.body.push(currentSubsection);
      }

      // Start a new subsection within the current section
      currentSubsection = {
        title: text,
        body: [],
      };
    } else if (currentSubsection) {
      // If there's an active subsection, add body to it
      currentSubsection.body.push(paragraph);
    } else if (currentSection) {
      // If there's no active subsection, add body to the current section
      currentSection.body.push(paragraph);
    } else {
      // If no section has started, add body to the preamble
      preamble.body.push(paragraph);
    }
  });

  // Push the last subsection and section if they exist
  if (currentSubsection) {
    currentSection.body.push(currentSubsection);
  }

  if (currentSection) {
    sections.push(currentSection);
  }

  // Include preamble if it has any body
  if (preamble.body.length > 0) {
    sections.unshift({
      title: "Cover Page",
      body: preamble.body,
    });
  }

  return sections;
}

// async function extractDocxContent(filePath) {
//   try {
//     // Step 1: Unzip the DOCX file and extract the document.xml and styles.xml
//     const outputDir = path.join(__dirname, "output");
//     await fs.promises.mkdir(outputDir, { recursive: true });

//     await fs
//       .createReadStream(filePath)
//       .pipe(unzipper.Extract({ path: outputDir }))
//       .promise();

//     // Load document.xml and styles.xml
//     const documentXmlPath = path.join(outputDir, "word", "document.xml");
//     const stylesXmlPath = path.join(outputDir, "word", "styles.xml");
//     const documentXml = await fs.promises.readFile(documentXmlPath, "utf8");
//     const stylesXml = await fs.promises.readFile(stylesXmlPath, "utf8");

//     // Use cheerio to parse XML files
//     const $doc = cheerio.load(documentXml, { xmlMode: true });
//     const $styles = cheerio.load(stylesXml, { xmlMode: true });

//     // Step 2: Create a mapping of styles from styles.xml
//     const styleMap = {};

//     $styles("w\\:style").each((_, style) => {
//       const styleId = $styles(style).attr("w:styleId");
//       const styleType = $styles(style).attr("w:type");

//       if (styleId && styleType) {
//         styleMap[styleId] = {
//           type: styleType,
//           name: $styles(style).find("w\\:name").attr("w:val"),
//           runProperties: extractRunStyles($styles(style).find("w\\:rPr")),
//           paragraphProperties: extractParagraphStyles(
//             $styles(style).find("w\\:pPr")
//           ),
//         };
//       }
//     });

//     // Function to extract run-level styles (inline styles like bold, italic, etc.)
//     function extractRunStyles(rPr) {
//       const styles = {};

//       if (rPr.find("w\\:b").length > 0) styles.bold = true;
//       if (rPr.find("w\\:i").length > 0) styles.italic = true;
//       if (rPr.find("w\\:u").length > 0) styles.underline = true;
//       if (rPr.find("w\\:strike").length > 0) styles.strikeThrough = true;

//       const color = rPr.find("w\\:color").attr("w:val");
//       if (color) styles.color = color;

//       const fontSize = rPr.find("w\\:sz").attr("w:val");
//       if (fontSize) styles.fontSize = fontSize;

//       const font = rPr.find("w\\:rFonts").attr("w:ascii");
//       if (font) styles.font = font;

//       const backgroundColor = rPr.find("w\\:shd").attr("w:fill");
//       if (backgroundColor) styles.backgroundColor = backgroundColor;

//       const highlight = rPr.find("w\\:highlight").attr("w:val");
//       if (highlight) styles.highlight = highlight;

//       return styles;
//     }

//     // Function to extract paragraph-level styles (block styles like alignment, spacing, etc.)
//     function extractParagraphStyles(pPr) {
//       const styles = {};

//       const alignment = pPr.find("w\\:jc").attr("w:val");
//       if (alignment) styles.alignment = alignment;

//       const spacingBefore = pPr.find("w\\:spacing").attr("w:before");
//       if (spacingBefore) styles.spacingBefore = spacingBefore;

//       const spacingAfter = pPr.find("w\\:spacing").attr("w:after");
//       if (spacingAfter) styles.spacingAfter = spacingAfter;

//       const indentLeft = pPr.find("w\\:ind").attr("w:left");
//       if (indentLeft) styles.indentLeft = indentLeft;

//       const indentRight = pPr.find("w\\:ind").attr("w:right");
//       if (indentRight) styles.indentRight = indentRight;

//       return styles;
//     }

//     // Step 5: Parse the document.xml content, applying styles from styles.xml
//     function parseElement(element) {
//       const children = [];

//       element.children().each((_, child) => {
//         const tag = $doc(child)[0].tagName;

//         if (tag === "w:p") {
//           // Handle paragraph
//           const paragraphData = {
//             type: "paragraph",
//             text: "",
//             styles: {},
//             children: [],
//           };

//           // Extract paragraph-level styles
//           const pPr = $doc(child).find("w\\:pPr");
//           const pStyleId = pPr.find("w\\:pStyle").attr("w:val");
//           // console.log(pStyleId);
//           if (pStyleId && styleMap[pStyleId]) {
//             // console.log(styleMap[pStyleId]);
//             paragraphData.styles = {
//               ...styleMap[pStyleId].paragraphProperties,
//               ...styleMap[pStyleId].runProperties,
//             };
//           }
//           if (pPr.length) {
//             paragraphData.styles = {
//               ...paragraphData.styles,
//               ...extractParagraphStyles(pPr),
//             };
//           }

//           // Extract run-level text and styles
//           $doc(child)
//             .find("w\\:r")
//             .each((_, run) => {
//               const runText = $doc(run).find("w\\:t").text();
//               const rPr = $doc(run).find("w\\:rPr");
//               const rStyleId = rPr.find("w\\:rStyle").attr("w:val");
//               let runStyles = {};

//               if (rStyleId && styleMap[rStyleId]) {
//                 runStyles = {
//                   ...styleMap[rStyleId].runProperties,
//                 };
//               }
//               runStyles = {
//                 ...runStyles,
//                 ...extractRunStyles(rPr),
//               };
//               paragraphData.children.push({ text: runText, styles: runStyles });
//             });

//           paragraphData.text = paragraphData.children
//             .map((child) => child.text)
//             .join("");
//           children.push(paragraphData);
//         } else if (tag === "w:tbl") {
//           // Handle table
//           const tableData = {
//             type: "table",
//             rows: [],
//           };

//           $doc(child)
//             .find("w\\:tr")
//             .each((_, row) => {
//               const rowData = [];

//               $doc(row)
//                 .find("w\\:tc")
//                 .each((_, cell) => {
//                   const cellData = {
//                     type: "cell",
//                     content: parseElement($doc(cell)),
//                   };
//                   rowData.push(cellData);
//                 });

//               tableData.rows.push(rowData);
//             });

//           children.push(tableData);
//         } else if (tag === "w:sectPr") {
//           // Handle section properties (if necessary)
//           const sectionData = {
//             type: "section",
//             styles: {
//               pageSize:
//                 $doc(child).find("w\\:pgSz").attr("w:w") +
//                 "x" +
//                 $doc(child).find("w\\:pgSz").attr("w:h"),
//               margins: {
//                 top: $doc(child).find("w\\:pgMar").attr("w:top"),
//                 bottom: $doc(child).find("w\\:pgMar").attr("w:bottom"),
//                 left: $doc(child).find("w\\:pgMar").attr("w:left"),
//                 right: $doc(child).find("w\\:pgMar").attr("w:right"),
//               },
//             },
//           };
//           children.push(sectionData);
//         } else {
//           // Handle other tags if necessary
//           children.push(parseElement($doc(child)));
//         }
//       });

//       return children.filter((child) =>
//         Array.isArray(child) ? child.length > 0 : true
//       );
//     }

//     // Step 6: Parse the document content
//     const documentContent = parseElement($doc("w\\:body"));

//     rimraf.sync(outputDir);

//     return documentContent;
//   } catch (error) {
//     console.error("Error processing DOCX file:", error);
//     return null;
//   }
// }

module.exports = { processDocxFile, extractDocxContent };

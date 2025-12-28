import parseXML from "./parseXML.js";
import ZIPExtractor from "./extractZIP.js";

const getChild = (parent, tag) => {
  if (!parent) return parent;
  return parent._children.find(c => c._tag === tag);
};

const parseStyle = (style) => {

  const textProperties = getChild(style, "style:text-properties");
  const paragraphProperties = getChild(style, "style:paragraph-properties");

  let css = "";

  if (textProperties) {

    let prop;
    if (prop = textProperties["fo:font-weight"]) {
      css += `font-weight:${prop};`;
    }
    if (prop = textProperties["fo:font-style"]) {
      css += `font-style:${prop};`;
    }
    if (prop = textProperties["style:text-underline-style"]) {
      // Hack, prevents colliding with line-through style
      css += `border-bottom:1px ${prop};`;
    }
    if (prop = textProperties["style:text-line-through-style"]) {
      css += `text-decoration:line-through ${prop};`;
    }
    if (prop = textProperties["fo:color"]) {
      css += `color:${prop};`;
    }
    if (prop = textProperties["fo:background-color"]) {
      css += `background-color:${prop};`;
    }
    if (prop = textProperties["fo:font-size"]) {
      css += `font-size:${prop};`;
    }
    if (prop = textProperties["style:font-name"]) {
      css += `font-family:${prop};`;
    }

  }

  if (paragraphProperties) {

    let prop;
    if (prop = paragraphProperties["fo:text-align"]) {
      css += `text-align:${prop};`;
    }

  }

  return css;

};

const handleElement = async (element, styles, zip, layout) => {

  const pageWidth = parseFloat(layout["fo:page-width"]);
  const pageHeight = parseFloat(layout["fo:page-height"]);

  let htmlTag = element._tag.split(":").pop();
  let attributes = "";
  let content = "";

  for (const child of element._children) {
    if (typeof child === "string") {
      content += child;
    } else {
      content += await handleElement(child, styles, zip, layout);
    }
  }

  const styleName = element[element._tag.split(":")[0] + ":style-name"];
  const style = styles.document.find(c => c["style:name"] === styleName);
  let css = parseStyle(style);

  switch (element._tag) {
    case "text:h": htmlTag = "h1"; break;
    case "text:list": htmlTag = "ul"; break;
    case "text:list-item": htmlTag = "li"; break;
    case "table:table":
      css += "width:100%;border-collapse:collapse";
      break;
    case "table:table-row": htmlTag = "tr"; break;
    case "table:table-cell":
      htmlTag = "td";
      css += "border:1px solid black;";
      break;
    case "draw:image":
    {
      htmlTag = "img";
      const mime = element["draw:mime-type"];
      const path = element["xlink:href"];
      const base64 = zip.extractBase64(path);
      attributes += ` src="data:${mime};base64,${base64}"`;
      css += "max-width:100%;";
      break;
    }
    case "draw:frame":
    {
      htmlTag = "div";
      const width = (parseFloat(element["svg:width"]) / pageWidth * 100).toFixed(3);
      const height = (parseFloat(element["svg:height"]) / pageHeight * 100).toFixed(3);
      const x = (parseFloat(element["svg:x"]) / pageWidth * 100).toFixed(3);
      const y = (parseFloat(element["svg:y"]) / pageHeight * 100).toFixed(3);
      css += `position:absolute;left:${x}%;top:${y}%;width:${width}%;height:${height}%;`;
      break;
    }
    default: break;
  }

  return `<${htmlTag}${css ? ` style="${css}"` : ""}${attributes}>${content}</${htmlTag}>`;

};

const extractDocument = async (bytes, callback) => {

  const zip = new ZIPExtractor(bytes);

  try {

    const documentXML = zip.extractText("content.xml");
    const stylesXML = zip.extractText("styles.xml");
    const document = parseXML(documentXML);
    const styles = parseXML(stylesXML);

    const styleGroups = {
      document: getChild(document[0], "office:automatic-styles")._children,
      main: getChild(styles[0], "office:styles")._children,
      automatic: getChild(styles[0], "office:automatic-styles")._children,
      master: getChild(styles[0], "office:master-styles")._children
    };

    const masterPage = styleGroups.master.find(c => c._tag === "style:master-page");
    const pageLayoutName = masterPage["style:page-layout-name"];
    const pageStyle = styleGroups.automatic.find(c => c._tag === "style:page-layout" && c["style:name"] === pageLayoutName);
    const pageLayoutProperties = getChild(pageStyle, "style:page-layout-properties");

    return await callback(zip, document, styleGroups, pageLayoutProperties);

  } catch (e) {

    console.error("Failed to parse ODP, using thumbnail instead.\n", e);

    const base64 = zip.extractBase64("Thumbnails/thumbnail.png");
    return `<img src="data:image/png;base64,${base64}" style="width:100%"></img>`;

  }

};

const parseODT = async (bytes) => {
  return await extractDocument(bytes, async (zip, doc, styles, layout) => {
    let outputHTML = "";

    const elements = getChild(getChild(doc[0], "office:body"), "office:text")._children;
    for (const element of elements) {
      outputHTML += await handleElement(element, styles, zip, layout);
    }

    return outputHTML;
  });
};

const parseODP = async (bytes) => {
  return await extractDocument(bytes, async (zip, doc, styles, layout) => {
    let outputHTML = `<style>.__page{position:relative;width:100%;aspect-ratio:16/9}</style>`;

    const pages = getChild(getChild(doc[0], "office:body"), "office:presentation")
      ._children.filter(c => c._tag === "draw:page");

    for (const page of pages) {
      let pageHTML = "";
      for (const element of page._children) {
        pageHTML += await handleElement(element, styles, zip, layout);
      }
      outputHTML += `<div class="__page">${pageHTML}</div>`;
    }

    return outputHTML;
  });
};

export { parseODT, parseODP };

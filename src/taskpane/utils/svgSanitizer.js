// Word insertion rasterises the rendered SVG through a canvas. As soon as that
// SVG contains a <foreignObject>, the browser marks the canvas as tainted and
// canvas.toDataURL() throws SecurityError, so the diagram can never be
// inserted. Mermaid 11 emits HTML labels for:
//
//   - venn-beta `text` nodes inside a set or union,
//   - event modeling cards,
//   - architecture-beta icons that carry `iconText`.
//
// The shared preview/export renderer runs this converter so the SVG handed to
// the canvas is rasterisable and preview uses the same labels as inserted PNG.
// Text is measured with a nominal character width: the goal is a readable label
// that stays inside its box, not pixel-identical HTML layout.

/* global DOMParser, XMLSerializer */

import { readSvgGeometry } from "./diagramUtils";

const SVG_NS = "http://www.w3.org/2000/svg";
const XLINK_NS = "http://www.w3.org/1999/xlink";

// Work with complete XML attributes: replacing `href="https://..."` in a
// string also matches the tail of `xlink:href`, leaving a broken `xlink:`.
// Kanban ticket links use precisely that namespaced form.
export function prepareSvgForRasterization(svg, rasterSize) {
  const withNamespaces = svg.replace(/<svg\b([^>]*)>/i, (tag, attributes) => {
    const xmlns = /\sxmlns\s*=/.test(attributes) ? "" : ` xmlns="${SVG_NS}"`;
    const xlink = /\sxmlns:xlink\s*=/.test(attributes) ? "" : ` xmlns:xlink="${XLINK_NS}"`;
    return `<svg${attributes}${xmlns}${xlink}>`;
  });
  const doc = new DOMParser().parseFromString(withNamespaces, "image/svg+xml");
  if (doc.getElementsByTagName("parsererror").length || doc.documentElement.localName !== "svg") {
    throw new Error("Invalid SVG XML: could not prepare the diagram image");
  }

  if (rasterSize) {
    const root = doc.documentElement;
    // Preserve the intrinsic coordinate system for callers without a viewBox.
    if (!root.hasAttribute("viewBox")) {
      const geometry = readSvgGeometry(root);
      root.setAttribute("viewBox", `0 0 ${geometry.width} ${geometry.height}`);
    }
    root.setAttribute("width", String(rasterSize.width));
    root.setAttribute("height", String(rasterSize.height));
    root.style.removeProperty("max-width");
    root.style.removeProperty("max-height");
    root.style.setProperty("width", `${rasterSize.width}px`);
    root.style.setProperty("height", `${rasterSize.height}px`);
  }

  for (const element of doc.getElementsByTagName("*")) {
    for (const attribute of Array.from(element.attributes)) {
      if (attribute.localName === "href" && /^(?:https?:)?\/\//i.test(attribute.value.trim())) {
        element.removeAttributeNode(attribute);
      }
    }
  }
  for (const style of doc.getElementsByTagName("style")) {
    style.textContent = style.textContent.replace(/@import[^;]*;/gi, "");
  }
  return new XMLSerializer().serializeToString(doc);
}

const DEFAULT_FONT_SIZE = 16;
// Arial averages roughly 0.55em per character across mixed case.
const CHAR_WIDTH_RATIO = 0.55;
const LINE_HEIGHT_RATIO = 1.2;
const BOX_PADDING_RATIO = 0.92;

function toNumber(value, fallback = 0) {
  const parsed = typeof value === "number" ? value : parseFloat(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function parseFontSize(element) {
  for (let node = element; node; node = node.parentNode) {
    const inline = node.getAttribute?.("style");
    const match = inline && /font-size\s*:\s*([\d.]+)px/i.exec(inline);
    if (match) {
      return toNumber(match[1], DEFAULT_FONT_SIZE);
    }
  }
  return DEFAULT_FONT_SIZE;
}

function wrapLabel(text, maxWidth, fontSize) {
  const maxChars = Math.max(Math.floor(maxWidth / (fontSize * CHAR_WIDTH_RATIO)), 4);
  const lines = [];

  for (const paragraph of text.split(/\r?\n/)) {
    const words = paragraph.trim().split(/\s+/).filter(Boolean);
    if (!words.length) {
      continue;
    }
    let current = words[0];
    for (const word of words.slice(1)) {
      if (current.length + 1 + word.length <= maxChars) {
        current += ` ${word}`;
      } else {
        lines.push(current);
        current = word;
      }
    }
    lines.push(current);
  }

  return lines.length ? lines : [""];
}

function labelClassFor(foreignObject) {
  // Mermaid puts the colour on a CSS class inside the HTML payload, e.g.
  // `.venn-text-node { fill: ... }`. Carrying that class over to <text> keeps
  // the theme colours without evaluating any CSS here.
  const span = foreignObject.querySelector("[class]");
  return span?.getAttribute("class") || foreignObject.getAttribute("class") || "";
}

function replaceWithText(foreignObject) {
  const doc = foreignObject.ownerDocument;
  const x = toNumber(foreignObject.getAttribute("x"));
  const y = toNumber(foreignObject.getAttribute("y"));
  const width = toNumber(foreignObject.getAttribute("width"));
  const height = toNumber(foreignObject.getAttribute("height"));
  const text = (foreignObject.textContent || "").replace(/\s+/g, " ").trim();
  const fontSize = parseFontSize(foreignObject);
  const centerX = x + width / 2;
  const centerY = y + height / 2;

  const textNode = doc.createElementNS(SVG_NS, "text");
  textNode.setAttribute("x", String(centerX));
  textNode.setAttribute("y", String(centerY));
  textNode.setAttribute("text-anchor", "middle");
  textNode.setAttribute("dominant-baseline", "central");
  textNode.setAttribute("font-size", String(fontSize));

  const className = labelClassFor(foreignObject);
  if (className) {
    textNode.setAttribute("class", className);
  }
  // Inherited colour from the surrounding group is the closest match for the
  // HTML label, and it keeps dark-theme labels readable on the white canvas.
  const inheritedFill = foreignObject.parentNode?.getAttribute?.("fill");
  if (inheritedFill && inheritedFill !== "none") {
    textNode.setAttribute("fill", inheritedFill);
  }

  const lines = width > 0 ? wrapLabel(text, width * BOX_PADDING_RATIO, fontSize) : [text];
  const lineHeight = fontSize * LINE_HEIGHT_RATIO;
  const firstLineY = centerY - ((lines.length - 1) * lineHeight) / 2;

  lines.forEach((line, index) => {
    const tspan = doc.createElementNS(SVG_NS, "tspan");
    tspan.setAttribute("x", String(centerX));
    tspan.setAttribute("y", String(firstLineY + index * lineHeight));
    tspan.textContent = line;
    textNode.appendChild(tspan);
  });

  foreignObject.parentNode.replaceChild(textNode, foreignObject);
}

/**
 * Replace every <foreignObject> with an equivalent <text> label.
 *
 * Returns the (possibly unchanged) SVG plus the number of replaced nodes and a
 * flag for a parse failure, so callers can tell "nothing to do" apart from
 * "could not inspect".
 */
export function stripForeignObjects(svg) {
  if (typeof svg !== "string" || !/<foreignObject/i.test(svg)) {
    return { svg, replaced: 0, failed: false };
  }

  let doc;
  try {
    doc = new DOMParser().parseFromString(svg, "image/svg+xml");
  } catch {
    return { svg, replaced: 0, failed: true };
  }

  if (!doc || doc.getElementsByTagName("parsererror").length) {
    return { svg, replaced: 0, failed: true };
  }

  const nodes = Array.from(doc.getElementsByTagName("foreignObject"));
  for (const node of nodes) {
    try {
      replaceWithText(node);
    } catch {
      // Leave this node alone; a single bad label must not abort the export.
    }
  }

  let result;
  try {
    result = new XMLSerializer().serializeToString(doc);
  } catch (error) {
    return { svg, replaced: 0, failed: true, error };
  }

  const remaining = (result.match(/<foreignObject/gi) || []).length;
  return { svg: result, replaced: nodes.length - remaining, failed: false };
}

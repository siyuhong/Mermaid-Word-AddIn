/* global document, window, DOMParser, XMLSerializer */

const PAINT_OPACITY = [
  ["fill", "fill-opacity"],
  ["stroke", "stroke-opacity"],
  ["stop-color", "stop-opacity"],
  ["flood-color", "flood-opacity"],
];

function splitColor(value) {
  const match = /^rgba?\(\s*([\d.]+)[, ]+([\d.]+)[, ]+([\d.]+)(?:\s*[,/]\s*([\d.]+))?\s*\)$/i.exec(
    value
  );
  if (!match) {
    return null;
  }
  return {
    color: `#${match
      .slice(1, 4)
      .map((channel) => Math.round(Number(channel)).toString(16).padStart(2, "0"))
      .join("")}`,
    alpha: match[4] === undefined ? 1 : Number(match[4]),
  };
}

function setPresentation(element, property, value) {
  element.setAttribute(property, value);
  // Inline priority also wins over Mermaid's !important theme declarations.
  element.style.setProperty(property, value, "important");
}

function materializeCssShadow(root, element, filter, index) {
  const match =
    /^drop-shadow\((rgba?\([^)]*\))\s+([-\d.]+)px\s+([-\d.]+)px(?:\s+([\d.]+)px)?\)$/.exec(filter);
  if (!match) return;
  const color = splitColor(match[1]);
  if (!color) return;
  const id = `${root.id || "mermaid"}-word-shadow-${index}`;
  const definition = root.ownerDocument.createElementNS(root.namespaceURI, "filter");
  for (const [key, value] of Object.entries({
    id,
    x: "-50%",
    y: "-50%",
    width: "200%",
    height: "200%",
    "color-interpolation-filters": "sRGB",
  })) {
    definition.setAttribute(key, value);
  }
  const shadow = root.ownerDocument.createElementNS(root.namespaceURI, "feDropShadow");
  for (const [key, value] of Object.entries({
    dx: match[2],
    dy: match[3],
    stdDeviation: match[4] || "0",
    "flood-color": color.color,
    "flood-opacity": color.alpha,
  })) {
    shadow.setAttribute(key, value);
  }
  definition.append(shadow);
  let defs = root.querySelector("defs");
  if (!defs) {
    defs = root.ownerDocument.createElementNS(root.namespaceURI, "defs");
    root.prepend(defs);
  }
  defs.append(definition);
  setPresentation(element, "filter", `url(#${id})`);
}

function expandDropShadows(root) {
  for (const [index, shadow] of [...root.querySelectorAll("feDropShadow")].entries()) {
    const input = shadow.getAttribute("in") || "SourceGraphic";
    const result = `word-shadow-${index}`;
    const primitive = (tag, attributes) => {
      const element = root.ownerDocument.createElementNS(root.namespaceURI, tag);
      for (const [key, value] of Object.entries(attributes)) element.setAttribute(key, value);
      return element;
    };
    // feDropShadow is SVG 2. Office loses the filtered shape. Its equivalent
    // SVG 1.1 primitives retain both the shadow and the original vector shape.
    const blur = primitive("feGaussianBlur", {
      in: input,
      stdDeviation: shadow.getAttribute("stdDeviation") || "2",
    });
    const offset = primitive("feOffset", {
      dx: shadow.getAttribute("dx") || "2",
      dy: shadow.getAttribute("dy") || "2",
      result,
    });
    const flood = primitive("feFlood", {
      "flood-color": shadow.getAttribute("flood-color") || "#000000",
      "flood-opacity": shadow.getAttribute("flood-opacity") || "1",
    });
    const composite = primitive("feComposite", { operator: "in", in2: result });
    const merge = primitive(
      "feMerge",
      shadow.hasAttribute("result") ? { result: shadow.getAttribute("result") } : {}
    );
    merge.append(primitive("feMergeNode", {}), primitive("feMergeNode", { in: input }));
    shadow.replaceWith(blur, offset, flood, composite, merge);
  }
}

function normalizeTextBaselines(root) {
  for (const text of root.querySelectorAll("text")) {
    const baseline = window.getComputedStyle(text).dominantBaseline;
    if (["auto", "alphabetic"].includes(baseline) || text.querySelector("tspan[y], tspan[dy]")) {
      continue;
    }
    const before = text.getBBox();
    setPresentation(text, "dominant-baseline", "alphabetic");
    const offset = before.y - text.getBBox().y;
    // Word treats hanging/middle/text-before-edge inconsistently, especially
    // with rotated axis labels. Resolve to the same measured alphabetic origin.
    // Keep transforms/clip paths in their original coordinate systems.
    if (Math.abs(offset) > 0.0001) {
      const positions = [...text.y.baseVal].map((length) => length.value + offset);
      text.setAttribute("y", positions.length ? positions.join(" ") : String(offset));
    }
  }
}

export function normalizeWordSvgStyles(svg) {
  const parsed = new DOMParser().parseFromString(svg, "image/svg+xml");
  const host = document.createElement("div");
  host.style.cssText = "position:fixed;left:-100000px;top:0;pointer-events:none";
  // Resolve CSS in isolation: app styles and a preview using the same SVG id
  // must not alter export colors. Do not inherit the measurement viewport's
  // opacity or visibility into the exported image.
  const shadow = host.attachShadow({ mode: "closed" });
  const root = document.importNode(parsed.documentElement, true);
  shadow.append(root);
  document.body.append(host);
  try {
    // Snapshot before changing parents, otherwise inherited alpha is applied
    // again on children. getComputedStyle resolves CSS variables/currentColor,
    // stylesheets, inline priority and inherited paint consistently with PNG.
    const styles = [root, ...root.querySelectorAll("*")]
      .filter((element) => element.style)
      .map((element) => {
        const computed = window.getComputedStyle(element);
        return {
          element,
          paints: PAINT_OPACITY.map(([paint, opacity]) => ({
            paint,
            opacity,
            color: splitColor(computed.getPropertyValue(paint)),
            alpha: parseFloat(computed.getPropertyValue(opacity)) || 0,
          })),
          filter: computed.filter,
        };
      });
    for (const [index, { element, paints, filter }] of styles.entries()) {
      for (const { paint, opacity, color, alpha } of paints) {
        if (color) {
          setPresentation(element, paint, color.color);
          setPresentation(element, opacity, String(color.alpha * alpha));
        }
      }
      materializeCssShadow(root, element, filter, index);
    }
    expandDropShadows(root);
    normalizeTextBaselines(root);
    return new XMLSerializer().serializeToString(root);
  } finally {
    host.remove();
  }
}

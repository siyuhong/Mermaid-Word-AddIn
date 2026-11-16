// Every diagram is rendered into one shared off-screen viewport. Mermaid
// derives the width of some diagrams (Gantt in particular) from the container
// it renders into, so a single measurement viewport makes the geometry
// deterministic for all diagram types without per-type width settings.
export const RENDER_VIEWPORT_WIDTH = 960;

const FALLBACK_GEOMETRY = { width: 800, height: 600 };

function toPositive(value) {
  const parsed = typeof value === "number" ? value : parseFloat(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 0;
}

/**
 * Read the intrinsic geometry of an SVG from either a string or a live element.
 * The viewBox is authoritative because Mermaid writes width="100%" for most
 * diagrams. Preview and Word insertion must both use this single reader.
 */
export function readSvgGeometry(source) {
  if (!source) {
    return { ...FALLBACK_GEOMETRY };
  }

  let viewBox;
  let width;
  let height;

  if (typeof source === "string") {
    const rootTag = source.match(/<svg\b[^>]*>/i)?.[0];
    if (!rootTag) {
      return { ...FALLBACK_GEOMETRY };
    }
    viewBox =
      rootTag.match(/\bviewBox\s*=\s*"([^"]*)"/i)?.[1] ??
      rootTag.match(/\bviewBox\s*=\s*'([^']*)'/i)?.[1];
    width =
      rootTag.match(/\swidth\s*=\s*"([^"]*)"/i)?.[1] ??
      rootTag.match(/\swidth\s*=\s*'([^']*)'/i)?.[1];
    height =
      rootTag.match(/\sheight\s*=\s*"([^"]*)"/i)?.[1] ??
      rootTag.match(/\sheight\s*=\s*'([^']*)'/i)?.[1];
  } else if (typeof source.getAttribute === "function") {
    viewBox = source.getAttribute("viewBox");
    width = source.getAttribute("width");
    height = source.getAttribute("height");
  } else {
    return { ...FALLBACK_GEOMETRY };
  }

  if (viewBox) {
    const values = String(viewBox)
      .trim()
      .split(/[\s,]+/)
      .map(parseFloat);
    if (values.length >= 4) {
      const boxWidth = toPositive(values[2]);
      const boxHeight = toPositive(values[3]);
      if (boxWidth && boxHeight) {
        return { width: boxWidth, height: boxHeight };
      }
    }
  }

  const attributeWidth = toPositive(width);
  const attributeHeight = toPositive(height);
  if (attributeWidth && attributeHeight) {
    return { width: attributeWidth, height: attributeHeight };
  }

  return { ...FALLBACK_GEOMETRY };
}

/**
 * Give every rendered SVG an explicit pixel size. Mermaid emits width="100%"
 * plus a max-width style for most diagrams, which makes <img> and canvas
 * sizing ambiguous. The original viewBox (including a non-zero origin) is kept.
 */
export function normalizeSvgSize(svgString) {
  const geometry = readSvgGeometry(svgString);
  if (typeof svgString !== "string" || !/<svg\b[^>]*>/i.test(svgString)) {
    return { svg: svgString, geometry };
  }

  const svg = svgString.replace(/<svg\b([^>]*)>/i, (tag, attributes) => {
    const cleaned = attributes
      .replace(/\s(?:width|height)\s*=\s*(?:"[^"]*"|'[^']*')/gi, "")
      .replace(/\sstyle\s*=\s*"([^"]*)"/gi, (styleMatch, styles) => {
        const kept = styles.replace(/max-width\s*:\s*[^;]+;?/gi, "").trim();
        return kept ? ` style="${kept}"` : "";
      });
    return `<svg${cleaned} width="${geometry.width}" height="${geometry.height}">`;
  });

  return { svg, geometry };
}

/**
 * Fit an intrinsic size into a box, preserving the aspect ratio exactly.
 * Preview passes pixels from its real container; Word insertion passes points
 * from the page setup. Neither caller needs to know the diagram type.
 */
export function fitIntoBox(geometry, maxBox) {
  const intrinsicWidth = toPositive(geometry?.width) || FALLBACK_GEOMETRY.width;
  const intrinsicHeight = toPositive(geometry?.height) || FALLBACK_GEOMETRY.height;
  const maxWidth = toPositive(maxBox?.width) || FALLBACK_GEOMETRY.width;
  const maxHeight = toPositive(maxBox?.height) || FALLBACK_GEOMETRY.height;
  const scale = Math.min(maxWidth / intrinsicWidth, maxHeight / intrinsicHeight);
  return { width: intrinsicWidth * scale, height: intrinsicHeight * scale };
}

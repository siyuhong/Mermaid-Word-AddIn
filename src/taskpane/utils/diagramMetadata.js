export const MERMAID_IMAGE_TITLE = "Mermaid Diagram";
const SOURCE_PREFIX = "MermaidSource:v1:";

export function encodeMermaidSource(code) {
  // Word normalizes literal tabs and CRLF in picture alternative text.
  // JSON escapes them so both PNG and SVG preserve the exact editor source.
  return SOURCE_PREFIX + JSON.stringify({ code });
}

export function decodeMermaidSource(description) {
  if (typeof description !== "string" || !description) return null;
  if (!description.startsWith("MermaidSource:")) {
    // Images from earlier versions stored raw Mermaid in alternative text.
    return description;
  }
  if (!description.startsWith(SOURCE_PREFIX)) return null;
  try {
    const data = JSON.parse(description.slice(SOURCE_PREFIX.length));
    return typeof data.code === "string" ? data.code : null;
  } catch {
    return null;
  }
}

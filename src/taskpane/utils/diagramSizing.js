import { fitIntoBox } from "./diagramUtils";

// Word uses points; Mermaid's viewBox uses CSS pixels (96 px per inch).
const POINTS_PER_INCH = 72;
const POINTS_PER_CSS_PIXEL = POINTS_PER_INCH / 96;
const TARGET_DPI = 600;
const MAX_CANVAS_PIXELS = 16000000;
const MAX_CANVAS_EDGE = 8192;
// Leave room for the inline picture's paragraph on a full-height page.
const INLINE_PARAGRAPH_SPACE_POINTS = 12;

export const DEFAULT_PAGE_DIMENSIONS = {
  pageWidth: 612,
  pageHeight: 792,
  marginLeft: 72,
  marginRight: 72,
  marginTop: 72,
  marginBottom: 72,
};

export function getDiagramInsertLayout(geometry, pageDimensions, options = {}) {
  const page = Object.fromEntries(
    Object.entries(DEFAULT_PAGE_DIMENSIONS).map(([key, fallback]) => {
      const value = pageDimensions?.[key];
      const valid = Number.isFinite(value) && (key.startsWith("margin") ? value >= 0 : value > 0);
      return [key, valid ? value : fallback];
    })
  );
  const box = {
    width: Math.max(1, page.pageWidth - page.marginLeft - page.marginRight),
    height: Math.max(
      1,
      page.pageHeight - page.marginTop - page.marginBottom - INLINE_PARAGRAPH_SPACE_POINTS
    ),
  };
  const naturalSize = {
    width: geometry.width * POINTS_PER_CSS_PIXEL,
    height: geometry.height * POINTS_PER_CSS_PIXEL,
  };
  const requestedScale = Number.parseFloat(options.scale);
  const scale =
    Number.isFinite(requestedScale) && requestedScale > 0
      ? Math.min(Math.max(requestedScale, 0.1), 10)
      : 1;
  // Apply zoom to the natural size first, then shrink only if either page
  // dimension is exceeded. Fitting before zoom would shrink large diagrams twice.
  const displaySize = fitIntoBox(
    { width: naturalSize.width * scale, height: naturalSize.height * scale },
    box,
    { allowUpscale: false }
  );

  // Export at the natural size, independent of preview zoom and page layout.
  // Word can enlarge a fitted picture later without losing these source pixels.
  const rawWidth = naturalSize.width * (TARGET_DPI / POINTS_PER_INCH);
  const rawHeight = naturalSize.height * (TARGET_DPI / POINTS_PER_INCH);
  const budgetScale = Math.min(
    1,
    Math.sqrt(MAX_CANVAS_PIXELS / (rawWidth * rawHeight)),
    MAX_CANVAS_EDGE / Math.max(rawWidth, rawHeight)
  );
  const rasterSize = {
    width: Math.max(1, Math.floor(rawWidth * budgetScale)),
    height: Math.max(1, Math.floor(rawHeight * budgetScale)),
  };
  return {
    displaySize,
    rasterSize,
    effectiveDpi: Math.round((rasterSize.width / displaySize.width) * POINTS_PER_INCH),
  };
}

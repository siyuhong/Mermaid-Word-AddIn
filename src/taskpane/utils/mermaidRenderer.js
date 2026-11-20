/* global document, console, DOMParser, XMLSerializer */
import mermaid from "mermaid";
import zenuml from "@mermaid-js/mermaid-zenuml";
import { RENDER_VIEWPORT_WIDTH, normalizeSvgSize } from "./diagramUtils";
import { stripForeignObjects } from "./svgSanitizer";
import { normalizeWordSvgStyles } from "./wordSvgStyles";

// Mermaid configuration is global. Keep initialization and rendering in the
// same queue so preview and export requests cannot change each other's theme.
let renderQueue = Promise.resolve();
let renderViewport = null;
let externalDiagramsReady = null;

function ensureExternalDiagrams() {
  if (!externalDiagramsReady) {
    // Register the detector once. Mermaid loads the ZenUML renderer on demand.
    externalDiagramsReady = mermaid.registerExternalDiagrams([zenuml]).catch((error) => {
      externalDiagramsReady = null;
      throw error;
    });
  }
  return externalDiagramsReady;
}

// One off-screen measurement viewport for every diagram type. It is laid out
// (not display:none) because Mermaid measures it while rendering and uses
// getBBox() on the SVG. The fixed width only stabilises layout; the final size
// always comes from the rendered viewBox plus fitIntoBox.
function getRenderViewport() {
  if (typeof document === "undefined" || !document.body) {
    return undefined;
  }
  if (!renderViewport || !document.body.contains(renderViewport)) {
    renderViewport = document.createElement("div");
    renderViewport.setAttribute("data-mermaid-measure-viewport", "");
    renderViewport.style.cssText = [
      "position:absolute",
      "left:-10000px",
      "top:0",
      `width:${RENDER_VIEWPORT_WIDTH}px`,
      "opacity:0",
      "pointer-events:none",
    ].join(";");
    document.body.appendChild(renderViewport);
  }
  return renderViewport;
}

function createExportCode(code) {
  const exportConfig = (config = {}) => ({
    ...config,
    htmlLabels: false,
    flowchart: { ...config.flowchart, htmlLabels: false },
    // Journey and timeline default to <foreignObject> text. Browsers mark a
    // canvas tainted when such an SVG is drawn into it, so canvas.toDataURL()
    // throws a SecurityError even for same-origin data. "old" keeps Mermaid's
    // coloured CSS classes and multi-line handling on plain <text>/<tspan>.
    journey: { ...config.journey, textPlacement: "old" },
    timeline: { ...config.timeline, textPlacement: "old" },
  });

  // Mindmap always renders `::icon(...)` decorations as <foreignObject>, even
  // though its labels are forced to plain <text>. Drawing that SVG into a
  // canvas taints it, so the icons must be removed before the export render.
  // They are not visible in the task pane either unless Font Awesome is loaded.
  const stripMindmapIcons = (source) => source.replace(/^[ \t]*::icon\([^)\r\n]*\)[ \t]*$/gm, "");

  // Match Mermaid's init directive syntax, including its single-quote support.
  // Preserve existing settings while overriding HTML labels and rasterisable
  // text placement in both the modern and legacy `config` shorthand.
  let exportCode = code.replace(
    /%%\{\s*(init|initialize)\s*:\s*([\s\S]*?)\}%%/gi,
    (directive, type, args) => {
      const config = JSON.parse(args.trim().replace(/'/g, '"'));
      const safeConfig = exportConfig(config);
      if (config.config) {
        safeConfig.config = exportConfig(config.config);
      }
      return `%%{${type}: ${JSON.stringify(safeConfig)}}%%`;
    }
  );

  // Append after the source so YAML frontmatter stays at the beginning and this
  // override takes precedence over frontmatter and any earlier init directives.
  exportCode = `${exportCode}\n%%{init: ${JSON.stringify(exportConfig())}}%%`;

  // detectType ignores frontmatter, directives and comments, so it is safe to
  // run on the transformed source and only affects actual mindmaps.
  if (mermaid.detectType(exportCode) === "mindmap") {
    exportCode = stripMindmapIcons(exportCode);
  }
  return exportCode;
}

// Mermaid reuses the box fill classes (`.task-type-N`, `.section-type-N`) for
// the plain `<text>` labels of journey/timeline. Without HTML labels those
// labels inherit the box colour and disappear. Restore the foreground colour
// that the preview gets from its HTML labels.
// Mermaid 12's use case shapes position SVG labels at the shape centre, but
// their stylesheet only centres HTML labels. Keep plain-text exports centred.
const EXPORT_TEXT_STYLE = `<style>
text.task, text.journey-section { fill: inherit !important; }
.usecase-element .label text { text-anchor: middle; }
</style>`;

function addExportTextStyles(svg) {
  return svg.replace(/<svg\b[^>]*>/i, (tag) => `${tag}${EXPORT_TEXT_STYLE}`);
}

function centerStateExportLabels(svg) {
  if (!svg.includes("statediagram-state")) {
    return svg;
  }
  const doc = new DOMParser().parseFromString(svg, "image/svg+xml");
  if (doc.querySelector("parsererror")) {
    return svg;
  }

  // Mermaid positions state labels at the left of their minimum-width box.
  // HTML centres the contents, but SVG text has no such box. Place each text
  // row at the node's centre in label coordinates, keeping its vertical offset.
  for (const label of doc.querySelectorAll(".statediagram-state > .label")) {
    // Only ordinary state boxes; notes and states with description sections
    // have their own text alignment and must keep it.
    if (!label.parentNode.querySelector(":scope > .basic.label-container")) {
      continue;
    }
    const transform = label.transform.baseVal.consolidate();
    const centerX = String(transform ? transform.matrix.inverse().e : 0);
    for (const text of label.querySelectorAll("text")) {
      text.setAttribute("text-anchor", "middle");
      text.setAttribute("x", centerX);
      for (const row of text.querySelectorAll(":scope > tspan")) {
        row.setAttribute("x", centerX);
      }
    }
  }
  return new XMLSerializer().serializeToString(doc);
}

export function renderMermaidSvg(id, code, config) {
  const request = renderQueue.then(async () => {
    await ensureExternalDiagrams();
    const isMindmap = mermaid.detectType(code, config) === "mindmap";
    mermaid.initialize({
      // Let mindmap choose its native cose-bilkent layout. Forcing ELK here
      // crashes Mermaid 12's mindmap renderer. Other diagrams default to ELK;
      // explicit caller settings, frontmatter and directives still take priority.
      ...(isMindmap ? {} : { layout: "elk" }),
      ...config,
      themeVariables: { ...config.themeVariables },
    });
    // Preview, PNG and SVG use one SVG-text layout. Rendering an HTML preview
    // and a different export layout changes wrapping, sizes and node positions.
    const renderCode = createExportCode(code);
    // Validate before rendering: Mermaid may leave an error SVG in the page
    // when render() receives invalid syntax, even with suppressErrorRendering.
    await mermaid.parse(renderCode);
    const result = await mermaid.render(id, renderCode, getRenderViewport());
    // Normalise once, here, so preview and Word insertion never guess a size.
    const { svg: sized, geometry } = normalizeSvgSize(result.svg);
    // The canvas that rasterises the export SVG must not see a foreignObject,
    // otherwise toDataURL() throws SecurityError and insertion is impossible.
    const { svg: sanitized, replaced, failed } = stripForeignObjects(sized);
    if (failed) {
      console.warn("Could not inspect the export SVG for foreignObject labels.");
    } else if (replaced) {
      console.log(`Converted ${replaced} foreignObject label(s) to <text> for export.`);
    }
    const compatibleSvg = normalizeWordSvgStyles(
      addExportTextStyles(centerStateExportLabels(sanitized))
    );
    return { ...result, svg: compatibleSvg, geometry };
  });

  // A failed request must not prevent subsequent diagrams from rendering.
  renderQueue = request.catch(() => {});
  return request;
}

export function validateMermaidCode(code) {
  const request = renderQueue.then(async () => {
    await ensureExternalDiagrams();
    return mermaid.parse(code);
  });

  renderQueue = request.catch(() => {});
  return request;
}

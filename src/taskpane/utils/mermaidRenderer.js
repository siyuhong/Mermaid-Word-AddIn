/* global document */
import mermaid from "mermaid";
import { RENDER_VIEWPORT_WIDTH, normalizeSvgSize } from "./diagramUtils";

// Mermaid configuration is global. Keep initialization and rendering in the
// same queue so preview and export requests cannot change each other's theme.
let renderQueue = Promise.resolve();
let renderViewport = null;

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
const EXPORT_TEXT_STYLE =
  "<style>text.task, text.journey-section { fill: inherit !important; }</style>";

function addExportTextStyles(svg) {
  return svg.replace(/<svg\b[^>]*>/i, (tag) => `${tag}${EXPORT_TEXT_STYLE}`);
}

export function renderMermaidSvg(id, code, config, forExport = false) {
  const request = renderQueue.then(async () => {
    mermaid.initialize({
      ...config,
      themeVariables: { ...config.themeVariables },
    });
    const renderCode = forExport ? createExportCode(code) : code;
    // Validate before rendering: Mermaid 10 may leave an error SVG in the page
    // when render() receives invalid syntax, even with suppressErrorRendering.
    await mermaid.parse(renderCode);
    const result = await mermaid.render(id, renderCode, getRenderViewport());
    // Normalise once, here, so preview and Word insertion never guess a size.
    const { svg, geometry } = normalizeSvgSize(result.svg);
    return { ...result, svg: forExport ? addExportTextStyles(svg) : svg, geometry };
  });

  // A failed request must not prevent subsequent diagrams from rendering.
  renderQueue = request.catch(() => {});
  return request;
}

export function validateMermaidCode(code) {
  const request = renderQueue.then(async () => {
    return mermaid.parse(code);
  });

  renderQueue = request.catch(() => {});
  return request;
}

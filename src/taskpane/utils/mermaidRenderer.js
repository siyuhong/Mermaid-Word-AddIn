import mermaid from "mermaid";

// Mermaid configuration is global. Keep initialization and rendering in the same
// queue so preview and export requests cannot change each other's theme.
let renderQueue = Promise.resolve();

function createExportCode(code) {
  // Match Mermaid's init directive syntax, including its single-quote support.
  // Preserve existing settings while overriding HTML labels in both the modern
  // flowchart config and the legacy diagram-specific `config` shorthand.
  const exportCode = code.replace(
    /%%\{\s*(init|initialize)\s*:\s*([\s\S]*?)\}%%/gi,
    (directive, type, args) => {
      const config = JSON.parse(args.trim().replace(/'/g, '"'));
      const safeConfig = {
        ...config,
        htmlLabels: false,
        flowchart: { ...config.flowchart, htmlLabels: false },
      };
      if (config.config) {
        safeConfig.config = { ...config.config, htmlLabels: false };
      }
      return `%%{${type}: ${JSON.stringify(safeConfig)}}%%`;
    }
  );

  // Append after the source so YAML frontmatter stays at the beginning and this
  // override takes precedence over frontmatter and any earlier init directives.
  return `${exportCode}\n%%{init: {"htmlLabels": false, "flowchart": {"htmlLabels": false}}}%%`;
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
    return mermaid.render(id, renderCode);
  });

  // A failed request must not prevent subsequent diagrams from rendering.
  renderQueue = request.catch(() => {});
  return request;
}

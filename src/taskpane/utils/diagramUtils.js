export const DIAGRAM_TYPES = {
  GANTT: "gantt",
  STATE: "stateDiagram",
  STATE_V2: "stateDiagram-v2",
  CLASS: "classDiagram",
};

export function detectDiagramType(code) {
  if (!code || typeof code !== "string") {
    return "";
  }

  const sanitizedLines = code
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line && !line.startsWith("%%"));

  for (const line of sanitizedLines) {
    const lowerLine = line.toLowerCase();

    if (lowerLine.startsWith("%%") || lowerLine.startsWith("---")) {
      continue;
    }

    if (lowerLine.startsWith("gantt")) {
      return DIAGRAM_TYPES.GANTT;
    }

    if (lowerLine.startsWith("state diagram")) {
      return DIAGRAM_TYPES.STATE;
    }

    if (lowerLine.startsWith("statediagram-v2")) {
      return DIAGRAM_TYPES.STATE_V2;
    }

    if (lowerLine.startsWith("statediagram")) {
      return DIAGRAM_TYPES.STATE;
    }

    if (lowerLine.startsWith("state ")) {
      return DIAGRAM_TYPES.STATE;
    }

    if (lowerLine.startsWith("flowchart") || lowerLine.startsWith("graph")) {
      return "flowchart";
    }

    if (lowerLine.startsWith("sequence")) {
      return "sequenceDiagram";
    }

    if (lowerLine.startsWith("classdiagram")) {
      return "classDiagram";
    }

    if (lowerLine.startsWith("class ")) {
      return "classDiagram";
    }

    if (lowerLine.startsWith("erdiagram")) {
      return "erDiagram";
    }

    if (lowerLine.startsWith("mindmap")) {
      return "mindmap";
    }

    if (lowerLine.startsWith("timeline")) {
      return "timeline";
    }

    if (lowerLine.startsWith("gitgraph")) {
      return "gitgraph";
    }

    if (lowerLine.startsWith("journey")) {
      return "journey";
    }

    if (lowerLine.startsWith("pie")) {
      return "pie";
    }

    if (lowerLine.startsWith("sankey")) {
      return "sankey";
    }

    if (lowerLine.startsWith("block")) {
      return "block";
    }
  }

  return "";
}

export function isStateDiagram(type) {
  return type === DIAGRAM_TYPES.STATE || type === DIAGRAM_TYPES.STATE_V2;
}

export function isStateDiagramV2(type) {
  return type === DIAGRAM_TYPES.STATE_V2;
}

export function isGanttDiagram(type) {
  return type === DIAGRAM_TYPES.GANTT;
}

export function isClassDiagram(type) {
  return type === DIAGRAM_TYPES.CLASS;
}

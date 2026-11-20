// Single source of truth for every diagram type the add-in knows about.
//
// Rendering itself always comes from Mermaid, so this file is not a capability
// list: it only feeds the editor (header detection, keyword colouring,
// completion) and documentation. Add a type here and the fallback lexer, the
// completion list and the docs stay in sync.
//
// `syntax` is the exact header keyword. `aliases` cover the forms Mermaid also
// accepts and the names `mermaid.detectType()` returns. `lezer` marks the types
// that codemirror-lang-mermaid parses with a real grammar; the rest are coloured
// by the shallow lexer in mermaidFallbackSyntax.js.

export const DIAGRAMS = [
  // --- Parsed by codemirror-lang-mermaid (Lezer grammars) -------------------
  { syntax: "flowchart", aliases: ["graph", "flowchart-v2"], label: "Flowchart", lezer: true },
  { syntax: "sequenceDiagram", label: "Sequence diagram", lezer: true },
  { syntax: "classDiagram", aliases: ["classDiagram-v2"], label: "Class diagram", lezer: true },
  {
    syntax: "stateDiagram-v2",
    aliases: ["stateDiagram"],
    label: "State diagram",
    lezer: true,
  },
  { syntax: "erDiagram", label: "Entity relationship diagram", lezer: true },
  { syntax: "gantt", label: "Gantt chart", lezer: true },
  { syntax: "pie", label: "Pie chart", lezer: true },
  { syntax: "journey", label: "User journey", lezer: true },
  { syntax: "requirementDiagram", label: "Requirement diagram", lezer: true },
  { syntax: "mindmap", label: "Mindmap", lezer: true },

  // --- Coloured by the fallback lexer, rendered by Mermaid -----------------
  { syntax: "gitGraph", label: "Git graph" },
  { syntax: "timeline", label: "Timeline" },
  { syntax: "quadrantChart", label: "Quadrant chart" },
  { syntax: "sankey-beta", label: "Sankey diagram" },
  { syntax: "block-beta", label: "Block diagram", keywords: ["columns", "block", "space"] },
  { syntax: "xychart-beta", label: "XY chart", keywords: ["x-axis", "y-axis", "bar", "line"] },
  {
    syntax: "C4Context",
    aliases: [
      "C4Container",
      "C4Component",
      "C4Dynamic",
      "C4Deployment",
      "Person",
      "Person_Ext",
      "System",
      "System_Ext",
      "SystemDb",
      "SystemDb_Ext",
      "SystemQueue",
      "SystemQueue_Ext",
      "Container",
      "Container_Ext",
      "ContainerDb",
      "ContainerDb_Ext",
      "ContainerQueue",
      "ContainerQueue_Ext",
      "Component",
      "Component_Ext",
      "Boundary",
      "Enterprise_Boundary",
      "System_Boundary",
      "Container_Boundary",
      "Deployment_Node",
      "Node",
      "Node_L",
      "Node_R",
      "Rel",
      "Rel_Back",
      "Rel_Neighbor",
      "BiRel",
      "Rel_Up",
      "Rel_Down",
      "Rel_Left",
      "Rel_Right",
      "Rel_U",
      "Rel_D",
      "Rel_L",
      "Rel_R",
      "UpdateRelStyle",
      "UpdateElementStyle",
      "UpdateLayoutConfig",
      "Lay_U",
      "Lay_D",
      "Lay_L",
      "Lay_R",
    ],
    label: "C4 diagram",
  },
  { syntax: "info", label: "Info" },
  {
    syntax: "swimlane-beta",
    aliases: ["swimlane"],
    label: "Swimlane diagram",
    keywords: ["subgraph", "end", "direction"],
  },

  // --- Mermaid 11/12 diagram types ----------------------------------------
  {
    syntax: "architecture-beta",
    aliases: ["architecture"],
    label: "Architecture diagram",
    keywords: ["group", "service", "junction", "in", "align", "row", "column"],
  },
  {
    syntax: "packet-beta",
    aliases: ["packet"],
    label: "Packet diagram",
    keywords: ["bitOrder"],
  },
  {
    syntax: "kanban",
    label: "Kanban board",
    // Cards are declared as `task1[Text]`; the words worth colouring are the
    // metadata keys, which are always standalone.
    keywords: ["ticket", "priority", "assigned", "until", "ticketBaseUrl"],
  },
  {
    syntax: "radar-beta",
    aliases: ["radar"],
    label: "Radar chart",
    keywords: ["axis", "curve", "showLegend", "ticks", "max", "min", "graticule"],
  },
  {
    syntax: "treemap-beta",
    aliases: ["treemap"],
    label: "Treemap",
    keywords: ["classDef", "class"],
  },
  {
    syntax: "treeView-beta",
    aliases: ["treeView"],
    label: "Tree view",
    keywords: ["title", "highlight", "annotation"],
  },
  {
    syntax: "venn-beta",
    aliases: ["venn"],
    label: "Venn diagram",
    keywords: ["set", "union", "text"],
  },
  {
    syntax: "wardley-beta",
    aliases: ["wardley"],
    label: "Wardley map",
    keywords: [
      "title",
      "size",
      "anchor",
      "component",
      "evolve",
      "note",
      "pipeline",
      "inertia",
      "label",
      "link",
    ],
  },
  {
    syntax: "ishikawa-beta",
    aliases: ["ishikawa"],
    label: "Ishikawa diagram",
    keywords: [],
  },
  {
    syntax: "cynefin-beta",
    aliases: ["cynefin"],
    label: "Cynefin framework",
    keywords: ["complex", "complicated", "clear", "chaotic", "confusion"],
  },
  {
    syntax: "eventmodeling",
    label: "Event modeling diagram",
    keywords: ["entity", "tf", "ui", "cmd", "evt", "data", "processor", "view", "note"],
  },
  {
    syntax: "railroad-beta",
    label: "Railroad diagram",
    keywords: [
      "title",
      "terminal",
      "nonterminal",
      "sequence",
      "choice",
      "optional",
      "oneOrMore",
      "zeroOrMore",
    ],
  },
  {
    syntax: "railroad-ebnf-beta",
    aliases: ["railroad-ebnf"],
    label: "Railroad diagram (EBNF)",
    keywords: ["title"],
  },
  {
    syntax: "railroad-abnf-beta",
    aliases: ["railroad-abnf"],
    label: "Railroad diagram (ABNF)",
    keywords: ["title"],
  },
  {
    syntax: "railroad-peg-beta",
    aliases: ["railroad-peg"],
    label: "Railroad diagram (PEG)",
    keywords: ["title"],
  },
  {
    syntax: "usecase-beta",
    aliases: ["usecase"],
    label: "Use case diagram",
    keywords: [
      "actor",
      "usecase",
      "systemBoundary",
      "direction",
      "stereotype",
      "include",
      "extend",
      "classDef",
      "class",
    ],
  },
  {
    // Registered by mermaidRenderer through @mermaid-js/mermaid-zenuml.
    syntax: "zenuml",
    label: "ZenUML sequence diagram",
    keywords: ["title", "if", "else", "par", "for", "while", "try", "catch", "return", "async"],
  },
];

// Statements shared by several types. Kept separate from the per-type lists so
// a word like `note` or `title` is only coloured once its type is detected.
export const SHARED_KEYWORDS = [
  "classDef",
  "linkStyle",
  "subgraph",
  "participant",
  "autonumber",
  "cherry-pick",
  "dateFormat",
  "axisFormat",
  "direction",
  "class",
  "style",
  "state",
  "note",
  "section",
  "title",
  "excludes",
  "includes",
  "commit",
  "branch",
  "checkout",
  "merge",
  "actor",
  "loop",
  "alt",
  "else",
  "opt",
  "par",
  "and",
  "rect",
  "end",
];

const syntaxIndex = new Map();

for (const diagram of DIAGRAMS) {
  const names = [diagram.syntax, ...(diagram.aliases || [])];
  for (const name of names) {
    syntaxIndex.set(name, diagram);
  }
}

/** Header keywords, longest first so `railroad-abnf-beta` beats `railroad-beta`. */
export const DIAGRAM_SYNTAXES = [...syntaxIndex.keys()].sort((a, b) => b.length - a.length);

export function getDiagram(syntax) {
  return syntaxIndex.get(syntax) || null;
}

/** Type-specific keywords for a header keyword, or an empty list. */
export function getKeywordsForSyntax(syntax) {
  const diagram = getDiagram(syntax);
  return diagram?.keywords || [];
}

const firstMeaningfulLine = (text) => {
  const lines = String(text || "").split(/\r?\n/);
  let index = 0;

  // Skip YAML frontmatter.
  if (lines[0]?.trim() === "---") {
    index = 1;
    while (index < lines.length && lines[index].trim() !== "---") {
      index += 1;
    }
    index += 1;
  }

  for (; index < lines.length; index += 1) {
    const line = lines[index].trim();
    if (!line || line.startsWith("%%")) {
      continue;
    }
    return line;
  }
  return "";
};

/**
 * Detect the diagram type from source text without invoking Mermaid.
 * Returns the header keyword Mermaid would see, or "" when unknown.
 */
export function detectSyntax(text) {
  const line = firstMeaningfulLine(text);
  if (!line) {
    return "";
  }
  const token = line.split(/[\s({[]/)[0];
  return syntaxIndex.has(token) ? token : "";
}

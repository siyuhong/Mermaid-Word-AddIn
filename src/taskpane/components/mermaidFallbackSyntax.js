// Lexical fallback for Mermaid diagram types that codemirror-lang-mermaid
// 0.5.0 does not have a Lezer grammar for. It is intentionally shallow: the
// goal is readable colour, not semantic validation.

export const DIAGRAM_TYPES = [
  "classDiagram",
  "classDiagram-v2",
  "stateDiagram-v2",
  "stateDiagram",
  "erDiagram",
  "gitGraph",
  "timeline",
  "mindmap",
  "pie",
  "gantt",
  "journey",
  "requirementDiagram",
  "flowchart",
  "graph",
  "sequenceDiagram",
  // Beta / experimental types have no Lezer grammar either. Listing them keeps
  // the diagram keyword coloured; rendering still depends on the Mermaid engine.
  "flowchart-elk",
  "sankey-beta",
  "block-beta",
  "quadrantChart",
  "xychart-beta",
  "C4Context",
  "C4Container",
  "C4Component",
  "C4Dynamic",
  "C4Deployment",
  "info",
];

export const KEYWORDS = [
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
  // C4 element and relationship keywords.
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
  // Block diagram statement keywords.
  "columns",
  "block",
  "space",
  // XY chart / quadrant chart axis and series keywords.
  "x-axis",
  "y-axis",
  "quadrant-1",
  "quadrant-2",
  "quadrant-3",
  "quadrant-4",
  "bar",
  "line",
];

// Ordered so that longer arrows win over their prefixes. ER endpoints
// (`||--o{`) include cardinality marks; flowchart edge labels are handled
// separately so `-->|Text|` does not turn `Text` into a node id.
export const ARROWS = [
  "||--o{",
  "||--|{",
  "||--o|",
  "}o--||",
  "}o--o{",
  "}o--|{",
  "|o--o{",
  "o--o{",
  "|--o{",
  "-->>",
  "..>>",
  "<|--",
  "<|..",
  "..|>",
  "<--",
  "-->",
  "->>",
  "--x",
  "--o",
  "o--",
  "*--",
  "..>",
  "-.->",
  "-.-",
  "---",
  "==>",
  "===",
  "->",
  "--",
  "==",
  "..",
];

const escapeRegExp = (value) => value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

const longestFirst = (values) =>
  [...new Set(values)].sort((a, b) => b.length - a.length).map(escapeRegExp);

const SIGNAL_PATTERN = new RegExp(
  `\\b(${longestFirst([...DIAGRAM_TYPES, ...KEYWORDS]).join("|")})\\b`,
  "g"
);
const ARROW_PATTERN = new RegExp(longestFirst(ARROWS).join("|"), "g");
const STRING_PATTERN = /"[^"\n]*"|'[^'\n]*'/g;
const COMMENT_PATTERN = /%%.*/g;
// Commas delimit Sankey fields and XY/quadrant series values, so a hyphenated
// name such as `Bio-conversion` must still be read as one id next to them.
const NODE_ID_PATTERN = /(^|[\s[(>{|,-])([A-Za-z_][\w-]*)(?=[\s[(<{|,)\]}:-]|$)/g;
const PLACEHOLDER_PATTERN = /[A-Za-z_][\w-]*/g;
const NUMBER_PATTERN = /\b\d+(?:\.\d+)?\b/g;
const PROPERTY_PATTERN = /(^|\s)([+\-#~]?\w+)(?=\s*:)/;
// `+String name` / `-eat() int` inside `class X { ... }`: colour the visibility
// modifier together with the type or method name, matching the `type-name`
// colour of the Lezer class grammar. The member name keeps the placeholder
// colour so attributes and methods read the same way.
const CLASS_MEMBER_PATTERN = /^([ \t]*)([+\-#~])([A-Za-z_]\w*)/;

const span = (from, to, className) => ({ from, to, className });

const collectMatches = (text, pattern, className) => {
  const spans = [];
  let match;
  pattern.lastIndex = 0;
  while ((match = pattern.exec(text))) {
    if (match[0].length > 0) {
      spans.push(span(match.index, match.index + match[0].length, className));
    }
  }
  return spans;
};

const overlaps = (spans, from, to) => spans.some((item) => from < item.to && to > item.from);

const pushIfFree = (spans, from, to, className) => {
  if (from >= to || overlaps(spans, from, to)) {
    return;
  }
  spans.push(span(from, to, className));
};

const maskSpans = (text, spans) => {
  // Keep UTF-16 units: every index produced by the RegExp scanners above is a
  // UTF-16 offset, which is also what CodeMirror positions use. Spreading the
  // string would collapse surrogate pairs and shift later spans on lines that
  // contain emoji or other astral characters.
  const chars = text.split("");
  for (const item of spans) {
    chars.fill(" ", item.from, item.to);
  }
  return chars.join("");
};

// Returns line-relative spans. `parsed` means a real Lezer sub-tree already
// covers the line; only comments and inline config are filled in then.
export function scanLine(text, parsed = false) {
  const spans = [];
  const leading = text.length - text.trimStart().length;
  const content = text.slice(leading);

  if (!content) {
    return spans;
  }
  if (content.startsWith("%%")) {
    return [span(leading, text.length, "comment")];
  }

  // Comments and inline directives take precedence over every other token.
  const commentSpans = collectMatches(content, COMMENT_PATTERN, "comment");
  for (const item of commentSpans) {
    pushIfFree(spans, leading + item.from, leading + item.to, item.className);
  }
  if (parsed) {
    return spans;
  }

  for (const item of collectMatches(content, STRING_PATTERN, "string")) {
    pushIfFree(spans, leading + item.from, leading + item.to, item.className);
  }

  SIGNAL_PATTERN.lastIndex = 0;
  const signal = SIGNAL_PATTERN.exec(content);
  if (signal) {
    pushIfFree(spans, leading + signal.index, leading + signal.index + signal[1].length, "keyword");
  }

  // Class members start with a visibility modifier (`+ - # ~`). Resolve them
  // before node ids because `-int` would otherwise match the `-` prefix rule in
  // NODE_ID_PATTERN and lose the modifier colour.
  const classMember = content.match(CLASS_MEMBER_PATTERN);
  if (classMember) {
    const typeStart = leading + classMember[1].length;
    pushIfFree(spans, typeStart, typeStart + 1 + classMember[3].length, "type-name");
  }

  // Never tokenise arrows or ids inside quoted strings.
  const masked = maskSpans(
    content,
    spans.map((item) => ({
      from: item.from - leading,
      to: item.to - leading,
    }))
  );

  for (const item of collectMatches(masked, ARROW_PATTERN, "arrow")) {
    pushIfFree(spans, leading + item.from, leading + item.to, item.className);
  }

  NODE_ID_PATTERN.lastIndex = 0;
  let idMatch;
  while ((idMatch = NODE_ID_PATTERN.exec(masked))) {
    const start = idMatch.index + idMatch[1].length;
    pushIfFree(spans, leading + start, leading + start + idMatch[2].length, "node-id");
  }

  // Numeric values (Sankey flows, quadrant coordinates, XY series data).
  for (const item of collectMatches(masked, NUMBER_PATTERN, "number")) {
    pushIfFree(spans, leading + item.from, leading + item.to, item.className);
  }

  // Remaining words such as style values (`fill:#f9f`) and edge labels still
  // need a readable colour on fallback diagrams.
  for (const item of collectMatches(masked, PLACEHOLDER_PATTERN, "placeholder")) {
    pushIfFree(spans, leading + item.from, leading + item.to, item.className);
  }

  const property = content.match(PROPERTY_PATTERN);
  if (property) {
    const start = leading + property.index + property[1].length;
    pushIfFree(spans, start, start + property[2].length, "node-id");
  }

  return spans;
}

export function scanDocument(doc, parsed = false) {
  const spans = [];
  let inFrontmatter = false;

  for (let index = 1; index <= doc.lines; index += 1) {
    const line = doc.line(index);
    const isFence = line.text.trim() === "---";
    if (index === 1 && isFence) {
      inFrontmatter = true;
    }

    if (inFrontmatter) {
      const leading = line.text.length - line.text.trimStart().length;
      spans.push(span(line.from + leading, line.to, "comment"));
      if (index > 1 && isFence) {
        inFrontmatter = false;
      }
    } else {
      for (const item of scanLine(line.text, parsed)) {
        spans.push(span(line.from + item.from, line.from + item.to, item.className));
      }
    }
  }

  return spans;
}

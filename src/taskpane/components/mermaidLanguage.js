import { LanguageSupport, syntaxTree } from "@codemirror/language";
import { parseMixed } from "@lezer/common";
import { tags as t, styleTags } from "@lezer/highlight";
import { Decoration, ViewPlugin } from "@codemirror/view";
import {
  flowchartLanguage,
  ganttLanguage,
  journeyLanguage,
  mermaidLanguage,
  mindmapLanguage,
  pieLanguage,
  requirementLanguage,
  sequenceLanguage,
} from "codemirror-lang-mermaid";
import { scanDocument } from "./mermaidFallbackSyntax";
import { DIAGRAMS, detectSyntax, getKeywordsForSyntax } from "../utils/diagramRegistry";

const FLOWCHART_EXTRA_STYLE_TAGS = {
  StyleKeyword: t.keyword,
  StyleText: t.string,
  NodeId: t.labelName,
  "⚠": t.invalid,
};

const patchedFlowchartLanguage = flowchartLanguage.configure({
  props: [styleTags(FLOWCHART_EXTRA_STYLE_TAGS)],
});

// Gantt parses dates, axis formats and task lines as plain `Text`, and mindmap
// leaves the `::icon(...)` keyword untagged, so neither ever receives a colour
// from the bundled grammars.
const GANTT_EXTRA_STYLE_TAGS = {
  Text: t.string,
};

const MINDMAP_EXTRA_STYLE_TAGS = {
  Icon: t.keyword,
};

const patchedGanttLanguage = ganttLanguage.configure({
  props: [styleTags(GANTT_EXTRA_STYLE_TAGS)],
});

const patchedMindmapLanguage = mindmapLanguage.configure({
  props: [styleTags(MINDMAP_EXTRA_STYLE_TAGS)],
});

// Rebuild parseMixed explicitly: LRLanguage.configure() replaces the parser
// without carrying over the previous wrapper mapping.
const patchedMermaidLanguage = mermaidLanguage.configure({
  wrap: parseMixed((node) => {
    switch (node.name) {
      case "MindmapDiagram":
        return { parser: patchedMindmapLanguage.parser };
      case "PieDiagram":
        return { parser: pieLanguage.parser };
      case "FlowchartDiagram":
        return { parser: patchedFlowchartLanguage.parser };
      case "SequenceDiagram":
        return { parser: sequenceLanguage.parser };
      case "JourneyDiagram":
        return { parser: journeyLanguage.parser };
      case "RequirementDiagram":
        return { parser: requirementLanguage.parser };
      case "GanttDiagram":
        return { parser: patchedGanttLanguage.parser };
      default:
        return null;
    }
  }),
});

// Statements and operators offered for every diagram type. Header keywords are
// generated from the registry at completion time, so a new Mermaid type shows
// up without touching this list.
const STATEMENT_KEYWORDS = [
  { label: "subgraph", type: "keyword", info: "Define a subgraph" },
  { label: "end", type: "keyword", info: "End a block" },
  { label: "linkStyle", type: "keyword", info: "Set link styles" },
  { label: "style", type: "keyword", info: "Define styles for nodes" },
  { label: "classDef", type: "keyword", info: "Define class styles" },
  { label: "applyClass", type: "keyword", info: "Apply class definitions" },
  { label: "columns", type: "keyword", info: "Set the column count (block-beta)" },
  { label: "x-axis", type: "keyword", info: "Define the x axis" },
  { label: "y-axis", type: "keyword", info: "Define the y axis" },
  { label: "quadrant-1", type: "keyword", info: "Set the top-right quadrant label" },
  { label: "quadrant-2", type: "keyword", info: "Set the top-left quadrant label" },
  { label: "quadrant-3", type: "keyword", info: "Set the bottom-left quadrant label" },
  { label: "quadrant-4", type: "keyword", info: "Set the bottom-right quadrant label" },
  { label: "bar", type: "keyword", info: "Draw bars (xychart-beta)" },
  { label: "line", type: "keyword", info: "Draw a line (xychart-beta)" },

  { label: "Person", type: "keyword", info: "Define a C4 person" },
  { label: "System", type: "keyword", info: "Define a C4 system" },
  { label: "SystemDb", type: "keyword", info: "Define a C4 system database" },
  { label: "SystemQueue", type: "keyword", info: "Define a C4 system queue" },
  { label: "Container", type: "keyword", info: "Define a C4 container" },
  { label: "ContainerDb", type: "keyword", info: "Define a C4 container database" },
  { label: "ContainerQueue", type: "keyword", info: "Define a C4 container queue" },
  { label: "Component", type: "keyword", info: "Define a C4 component" },
  { label: "Boundary", type: "keyword", info: "Define a C4 boundary" },
  { label: "System_Boundary", type: "keyword", info: "Define a C4 system boundary" },
  { label: "Container_Boundary", type: "keyword", info: "Define a C4 container boundary" },
  { label: "Deployment_Node", type: "keyword", info: "Define a C4 deployment node" },
  { label: "Rel", type: "keyword", info: "Define a C4 relationship" },
  { label: "BiRel", type: "keyword", info: "Define a bidirectional C4 relationship" },
  { label: "UpdateLayoutConfig", type: "keyword", info: "Configure C4 layout" },
  { label: "UpdateElementStyle", type: "keyword", info: "Configure C4 element style" },
  { label: "UpdateRelStyle", type: "keyword", info: "Configure C4 relationship style" },

  { label: "class", type: "keyword", info: "Define a class" },
  { label: "interface", type: "keyword", info: "Define an interface" },
  { label: "enum", type: "keyword", info: "Define an enumeration" },
  { label: "namespace", type: "keyword", info: "Define a namespace" },
  { label: "<|--", type: "relation", info: "Inheritance" },
  { label: "*--", type: "relation", info: "Composition" },
  { label: "o--", type: "relation", info: "Aggregation" },
  { label: "-->", type: "relation", info: "Association" },
  { label: "--", type: "relation", info: "Link (solid)" },
  { label: "..>", type: "relation", info: "Dependency" },
  { label: "..|>", type: "relation", info: "Realization" },

  { label: "state", type: "keyword", info: "Define a state" },
  { label: "[*]", type: "keyword", info: "Initial/final state" },
  { label: "note", type: "keyword", info: "Add a note" },
  { label: "right of", type: "position", info: "Position to the right" },
  { label: "left of", type: "position", info: "Position to the left" },
  { label: "above", type: "position", info: "Position above" },
  { label: "below", type: "position", info: "Position below" },

  { label: "title", type: "keyword", info: "Set the title" },
  { label: "dateFormat", type: "keyword", info: "Set the date format" },
  { label: "section", type: "keyword", info: "Define a section" },
  { label: "includes", type: "keyword", info: "Include tasks" },
  { label: "excludes", type: "keyword", info: "Exclude tasks" },

  { label: "-->", type: "arrow", info: "Arrow with filled head" },
  { label: "->>", type: "arrow", info: "Arrow with open head" },
  { label: "--x", type: "arrow", info: "Arrow with cross" },
  { label: "-x", type: "arrow", info: "Dashed arrow with cross" },
  { label: "---", type: "arrow", info: "Dashed line" },
];

// Header completion entries, built from the registry so the list always matches
// what the bundled Mermaid engine can actually render.
const SYNTAX_KEYWORDS = DIAGRAMS.map((diagram) => ({
  label: diagram.syntax,
  type: "keyword",
  info: diagram.label,
}));

// Aliases are kept out of the registry-derived list (C4 alone would add forty
// element names), but `graph` is common enough to deserve an entry.
const EXTRA_HEADER_KEYWORDS = [
  { label: "graph", type: "keyword", info: "Flowchart (graph alias)" },
];

const mergeOptions = (groups) => {
  const seen = new Set();
  const merged = [];
  for (const option of groups.flat()) {
    if (seen.has(option.label)) {
      continue;
    }
    seen.add(option.label);
    merged.push(option);
  }
  return merged;
};

export function mermaidCompletion(context) {
  const word = context.matchBefore(/\w*/);
  if (word.from === word.to && !context.explicit) {
    return null;
  }

  const syntax = detectSyntax(context.state.doc.toString());
  // Once the header names a type, its own statements are worth offering too;
  // before that only generic statements make sense.
  const typedKeywords = syntax
    ? getKeywordsForSyntax(syntax).map((label) => ({
        label,
        type: "keyword",
        info: `${syntax} statement`,
      }))
    : [];

  const options = mergeOptions([
    SYNTAX_KEYWORDS,
    EXTRA_HEADER_KEYWORDS,
    typedKeywords,
    STATEMENT_KEYWORDS,
  ]);
  const prefix = word.text.toLowerCase();

  return {
    from: word.from,
    options: options.filter((keyword) => keyword.label.toLowerCase().startsWith(prefix)),
  };
}

const hasUnparsedDiagram = (state) => {
  const tree = syntaxTree(state);
  if (tree.topNode.name !== "MermaidDiagram") {
    return false;
  }
  const children = [];
  for (let child = tree.topNode.firstChild; child; child = child.nextSibling) {
    children.push(child.name);
  }
  return children.every((name) => name === "⚠");
};

const collectTokenRanges = (state) => {
  const ranges = [];
  syntaxTree(state).iterate({
    enter(node) {
      if (!node.node.firstChild && node.to > node.from && node.name !== "⚠") {
        ranges.push([node.from, node.to]);
      }
      return true;
    },
  });
  ranges.sort((left, right) => left[0] - right[0]);
  return ranges;
};

const DECORATIONS = {
  keyword: Decoration.mark({ class: "cm-mermaid-keyword" }),
  string: Decoration.mark({ class: "cm-mermaid-string" }),
  arrow: Decoration.mark({ class: "cm-mermaid-arrow" }),
  comment: Decoration.mark({ class: "cm-mermaid-comment" }),
  "node-id": Decoration.mark({ class: "cm-mermaid-node-id" }),
  placeholder: Decoration.mark({ class: "cm-mermaid-placeholder" }),
  "type-name": Decoration.mark({ class: "cm-mermaid-type-name" }),
  number: Decoration.mark({ class: "cm-mermaid-number" }),
};

// Leaf tokens never overlap, so the last range starting before `to` is the only
// candidate that can intersect [from, to).
const overlapsToken = (ranges, from, to) => {
  let low = 0;
  let high = ranges.length - 1;
  let candidate = -1;
  while (low <= high) {
    const middle = (low + high) >> 1;
    if (ranges[middle][0] < to) {
      candidate = middle;
      low = middle + 1;
    } else {
      high = middle - 1;
    }
  }
  return candidate >= 0 && ranges[candidate][1] > from;
};

const buildDecorations = (state) => {
  const unparsed = hasUnparsedDiagram(state);
  const occupied = collectTokenRanges(state);
  const candidates = [];
  // Parsed diagrams already carry node and string colours from the grammar, so
  // only comments, directives and frontmatter need scanning there. That also
  // keeps the per-keystroke cost low for the diagrams with a Lezer grammar.
  const scanned = scanDocument(state.doc, !unparsed);

  for (const item of scanned) {
    const from = item.from;
    const to = item.to;
    if (from >= to || overlapsToken(occupied, from, to)) {
      continue;
    }
    candidates.push([from, to, item.className]);
  }

  if (!candidates.length) {
    return Decoration.none;
  }

  candidates.sort((left, right) => left[0] - right[0] || left[1] - right[1]);
  return Decoration.set(
    candidates.map(([from, to, className]) => DECORATIONS[className].range(from, to))
  );
};

const EDGE_LABEL_CLASS = "cm-mermaid-edge-label";

const edgeLabelDecorations = ViewPlugin.fromClass(
  class {
    constructor(view) {
      this.tree = syntaxTree(view.state);
      this.decorations = this.build(view);
    }

    update(update) {
      const tree = syntaxTree(update.state);
      if (update.docChanged || update.viewportChanged || tree !== this.tree) {
        this.tree = tree;
        this.decorations = this.build(update.view);
      }
    }

    build(view) {
      const ranges = [];
      const tree = syntaxTree(view.state);
      const doc = view.state.doc;
      tree.iterate({
        enter(node) {
          if (node.name !== "NodeText" || !node.node.parent) {
            return true;
          }
          const parent = node.node.parent;
          if (parent.name !== "Node") {
            return true;
          }
          const parentText = doc.sliceString(parent.from, parent.to);
          if (!parentText.trimStart().startsWith("|")) {
            return true;
          }
          ranges.push({ from: node.from, to: node.to });
          return true;
        },
      });
      return ranges.length
        ? Decoration.set(
            ranges.map(({ from, to }) =>
              Decoration.mark({ class: EDGE_LABEL_CLASS }).range(from, to)
            ),
            true
          )
        : Decoration.none;
    }
  },
  { decorations: (plugin) => plugin.decorations }
);

const fallbackDecorations = ViewPlugin.fromClass(
  class {
    constructor(view) {
      this.tree = syntaxTree(view.state);
      this.decorations = buildDecorations(view.state);
    }

    update(update) {
      const tree = syntaxTree(update.state);
      if (update.docChanged || update.viewportChanged || tree !== this.tree) {
        this.tree = tree;
        this.decorations = buildDecorations(update.state);
      }
    }
  },
  { decorations: (plugin) => plugin.decorations }
);

export default [
  new LanguageSupport(patchedMermaidLanguage),
  edgeLabelDecorations,
  fallbackDecorations,
];

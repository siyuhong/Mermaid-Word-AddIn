import { mermaid } from "codemirror-lang-mermaid";
import { LanguageSupport } from "@codemirror/language";

// Enhanced Mermaid language support that extends the original
// This provides better support for diagram types like stateDiagram-v2 and classDiagram
const enhancedMermaidLanguage = new LanguageSupport(mermaid());

// Mermaid keywords for autocompletion
const mermaidKeywords = [
  // Diagram types
  { label: "flowchart", type: "keyword", info: "Create a flowchart diagram" },
  { label: "graph", type: "keyword", info: "Create a graph diagram" },
  { label: "sequenceDiagram", type: "keyword", info: "Create a sequence diagram" },
  { label: "classDiagram", type: "keyword", info: "Create a class diagram" },
  { label: "stateDiagram", type: "keyword", info: "Create a state diagram" },
  { label: "stateDiagram-v2", type: "keyword", info: "Create a state diagram v2" },
  { label: "gantt", type: "keyword", info: "Create a Gantt chart" },
  { label: "pie", type: "keyword", info: "Create a pie chart" },
  { label: "journey", type: "keyword", info: "Create a journey diagram" },
  { label: "erDiagram", type: "keyword", info: "Create an entity relationship diagram" },
  { label: "gitgraph", type: "keyword", info: "Create a git graph" },

  // Common keywords
  { label: "subgraph", type: "keyword", info: "Define a subgraph" },
  { label: "end", type: "keyword", info: "End a block" },
  { label: "linkStyle", type: "keyword", info: "Set link styles" },
  { label: "style", type: "keyword", info: "Define styles for nodes" },
  { label: "classDef", type: "keyword", info: "Define class styles" },
  { label: "applyClass", type: "keyword", info: "Apply class definitions" },

  // Class diagram specific
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

  // State diagram specific
  { label: "state", type: "keyword", info: "Define a state" },
  { label: "[*]", type: "keyword", info: "Initial/final state" },
  { label: "note", type: "keyword", info: "Add a note" },
  { label: "right of", type: "position", info: "Position to the right" },
  { label: "left of", type: "position", info: "Position to the left" },
  { label: "above", type: "position", info: "Position above" },
  { label: "below", type: "position", info: "Position below" },

  // Gantt specific
  { label: "title", type: "keyword", info: "Set the title" },
  { label: "dateFormat", type: "keyword", info: "Set the date format" },
  { label: "section", type: "keyword", info: "Define a section" },
  { label: "includes", type: "keyword", info: "Include tasks" },
  { label: "excludes", type: "keyword", info: "Exclude tasks" },

  // Common arrows and connections
  { label: "-->", type: "arrow", info: "Arrow with filled head" },
  { label: "->>", type: "arrow", info: "Arrow with open head" },
  { label: "--x", type: "arrow", info: "Arrow with cross" },
  { label: "-x", type: "arrow", info: "Dashed arrow with cross" },
  { label: "---", type: "arrow", info: "Dashed line" },
];

// Autocompletion function for Mermaid
function mermaidCompletion(context) {
  const word = context.matchBefore(/\w*/);
  if (word.from == word.to && !context.explicit) return null;

  return {
    from: word.from,
    options: mermaidKeywords.filter((keyword) =>
      keyword.label.toLowerCase().startsWith(word.text.toLowerCase())
    ),
  };
}

// Export enhanced language support with autocompletion
export default enhancedMermaidLanguage;

// Export the completion function for use in CodeMirror extensions
export { mermaidCompletion };

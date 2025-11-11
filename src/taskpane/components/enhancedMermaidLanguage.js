import { mermaid } from "codemirror-lang-mermaid";

// Enhanced Mermaid language support that extends the original
// This provides better support for diagram types like stateDiagram-v2 and classDiagram
// Call the mermaid() function to get the LanguageSupport instance
const enhancedMermaidLanguage = mermaid();

export default enhancedMermaidLanguage;

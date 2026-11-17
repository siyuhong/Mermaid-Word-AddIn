import { HighlightStyle } from "@codemirror/language";
import { EditorView } from "@codemirror/view";
import { tags as t } from "@lezer/highlight";
import {
  flowchartTags,
  ganttTags,
  journeyTags,
  mindmapTags,
  pieTags,
  requirementTags,
  sequenceTags,
} from "codemirror-lang-mermaid";
import { tokens, webDarkTheme } from "@fluentui/react-components";

// A dedicated palette for the editor. Every tag gets an explicit colour so
// node identifiers, node labels, edge labels and comments stay distinguishable.
// Fluent token values are CSS custom properties, so the light/dark variants
// only differ in which palette object provides the fallback colours.
const createHighlightSpecs = (palette) => [
  {
    tag: [t.typeName, pieTags.diagramName, mindmapTags.diagramName],
    color: palette.purple,
    fontWeight: "600",
  },
  {
    tag: [
      t.keyword,
      t.controlKeyword,
      flowchartTags.orientation,
      pieTags.showData,
      pieTags.title,
      ganttTags.keyword,
      sequenceTags.keyword1,
      sequenceTags.keyword2,
      journeyTags.keyword,
      requirementTags.keyword,
    ],
    color: palette.brand,
    fontWeight: "600",
  },
  // Node labels: the text inside a shape.
  {
    tag: [flowchartTags.nodeText, journeyTags.text, pieTags.titleText],
    color: palette.foreground1,
    fontWeight: "500",
  },
  // Message text in sequence diagrams (flowchart edge labels share the node
  // text tag upstream, so they are distinguished by a ViewPlugin).
  {
    tag: [sequenceTags.messageText1, sequenceTags.messageText2],
    color: palette.foreground2,
    fontStyle: "italic",
  },
  // Identifier-like tokens: node ids, actors, class names, requirement ids.
  {
    tag: [
      flowchartTags.nodeId,
      sequenceTags.nodeText,
      journeyTags.actor,
      requirementTags.unquotedString,
      t.variableName,
      t.labelName,
    ],
    color: palette.blue,
  },
  // String literals stay distinct from both node and edge labels.
  {
    tag: [
      t.string,
      flowchartTags.string,
      pieTags.string,
      ganttTags.string,
      requirementTags.quotedString,
    ],
    color: palette.marigold,
  },
  // Mindmap levels get progressively cooler colours so hierarchy is visible.
  { tag: mindmapTags.lineText1, color: palette.foreground1 },
  { tag: mindmapTags.lineText2, color: palette.foreground2 },
  { tag: mindmapTags.lineText3, color: palette.blue },
  { tag: mindmapTags.lineText4, color: palette.teal },
  { tag: mindmapTags.lineText5, color: palette.purple },
  {
    tag: [t.number, t.atom, pieTags.number, journeyTags.score],
    color: palette.teal,
  },
  {
    tag: [
      t.comment,
      t.lineComment,
      t.blockComment,
      flowchartTags.lineComment,
      sequenceTags.lineComment,
      journeyTags.lineComment,
      requirementTags.lineComment,
      pieTags.lineComment,
      ganttTags.lineComment,
    ],
    color: palette.foreground4,
    fontStyle: "italic",
  },
  {
    tag: [
      t.contentSeparator,
      t.operator,
      flowchartTags.link,
      flowchartTags.nodeEdge,
      sequenceTags.arrow,
      requirementTags.arrow,
    ],
    color: palette.foreground3,
    fontWeight: "600",
  },
  { tag: t.invalid, color: palette.red, textDecoration: "wavy underline" },
];

const lightPalette = {
  brand: tokens.colorBrandForeground1,
  foreground1: tokens.colorNeutralForeground1,
  foreground2: tokens.colorNeutralForeground2,
  foreground3: tokens.colorNeutralForeground3,
  foreground4: tokens.colorNeutralForeground4,
  blue: tokens.colorPaletteBlueForeground2,
  marigold: tokens.colorPaletteMarigoldForeground2,
  teal: tokens.colorPaletteTealForeground2,
  purple: tokens.colorPalettePurpleForeground2,
  red: tokens.colorPaletteRedForeground1,
};

const darkPalette = {
  brand: webDarkTheme.colorBrandForeground1,
  foreground1: webDarkTheme.colorNeutralForeground1,
  foreground2: webDarkTheme.colorNeutralForeground2,
  foreground3: webDarkTheme.colorNeutralForeground3,
  foreground4: webDarkTheme.colorNeutralForeground4,
  blue: webDarkTheme.colorPaletteBlueForeground2,
  marigold: webDarkTheme.colorPaletteMarigoldForeground2,
  teal: webDarkTheme.colorPaletteTealForeground2,
  purple: webDarkTheme.colorPalettePurpleForeground2,
  red: webDarkTheme.colorPaletteRedForeground1,
};

export const mermaidHighlightStyle = HighlightStyle.define(createHighlightSpecs(lightPalette));

export const mermaidHighlightStyleDark = HighlightStyle.define(createHighlightSpecs(darkPalette));

const createEditorThemeSpec = (theme) => ({
  "&": {
    backgroundColor: theme.colorNeutralBackground1,
    color: theme.colorNeutralForeground1,
    fontSize: "13px",
    fontFamily: tokens.fontFamilyMonospace,
  },
  ".cm-content": { caretColor: theme.colorBrandForeground1 },
  ".cm-gutters": {
    backgroundColor: theme.colorNeutralBackground2,
    color: theme.colorNeutralForeground3,
    borderRight: `1px solid ${theme.colorNeutralStroke2}`,
  },
  // Keep CodeMirror's translucent active-line background: an opaque fill
  // would hide the selection layer rendered underneath the text.
  ".cm-activeLineGutter": { backgroundColor: theme.colorNeutralBackground3 },
  ".cm-selectionBackground, &.cm-focused .cm-selectionBackground": {
    backgroundColor: theme.colorNeutralBackground1Selected,
  },
  ".cm-lintRange-error": {
    backgroundImage: "none",
    textDecoration: "underline wavy",
    textDecorationColor: theme.colorPaletteRedForeground1,
  },
  ".cm-mermaid-keyword": {
    color: theme.colorBrandForeground1,
    fontWeight: "600",
  },
  ".cm-mermaid-string": { color: theme.colorPaletteMarigoldForeground2 },
  ".cm-mermaid-arrow": {
    color: theme.colorNeutralForeground3,
    fontWeight: "600",
  },
  ".cm-mermaid-comment": {
    color: theme.colorNeutralForeground4,
    fontStyle: "italic",
  },
  ".cm-mermaid-node-id": { color: theme.colorPaletteBlueForeground2 },
  // Numbers and free text keep the same colours as the Lezer-driven diagrams
  // (`t.number` is teal, node labels use the primary foreground) so fallback
  // diagrams read like the fully parsed ones.
  ".cm-mermaid-number": { color: theme.colorPaletteTealForeground2 },
  ".cm-mermaid-placeholder": { color: theme.colorNeutralForeground1 },
  ".cm-mermaid-edge-label": {
    color: theme.colorNeutralForeground2,
    fontStyle: "italic",
  },
  ".cm-mermaid-type-name": {
    color: theme.colorPalettePurpleForeground2,
    fontWeight: "600",
  },
});

export const mermaidEditorTheme = EditorView.theme(createEditorThemeSpec(tokens));
export const mermaidEditorThemeDark = EditorView.theme(createEditorThemeSpec(webDarkTheme), {
  dark: true,
});

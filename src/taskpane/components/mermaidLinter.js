import { linter } from "@codemirror/lint";
import { validateMermaidCode } from "../utils/mermaidRenderer";
import { getMermaidSourcePosition } from "../utils/mermaidSourceMap";

const MESSAGE_LIMIT = 220;

const getErrorMessage = (error) => {
  if (!error) {
    return "Mermaid syntax error";
  }
  if (typeof error === "string") {
    return error;
  }
  return error.str || error.message || error.error?.message || "Mermaid syntax error";
};

const getErrorLocation = (error) => {
  const hash = error?.hash || error?.error?.hash;
  const loc = hash?.loc || error?.loc;
  if (!loc) {
    return null;
  }

  const line = Number(loc.first_line);
  const column = Number(loc.first_column);

  // Mermaid's flowchart parser reports a location from the offending token to
  // end-of-input, so `first_line` is the real start while `last_line` often
  // points past the document. Other parsers report a single line. Either way
  // only the start matters for the underline.
  const hasLine = Number.isFinite(line) && line > 0;
  const hasColumn = Number.isFinite(column) && column >= 0;
  if (!hasLine && !hasColumn) {
    return null;
  }

  return {
    line: hasLine ? line : 1,
    column: hasColumn ? column : 0,
  };
};

const toDiagnostic = (state, error) => {
  const doc = state.doc || state;
  const location = getErrorLocation(error);
  let from = 0;
  let to = 0;

  if (location) {
    const position = getMermaidSourcePosition(doc.toString(), location.line, location.column);
    const line = doc.lineAt(position);
    // Underline a visible character on this line even when the parser points
    // at its end, instead of marking the newline or the following line.
    from = Math.min(position, Math.max(line.from, line.to - 1));
    to = line.to > from ? line.to : Math.min(from + 1, doc.length);
  }

  if (to <= from) {
    to = Math.min(from + 1, doc.length);
  }

  // A zero-length diagnostic cannot be rendered by CodeMirror.
  if (to <= from && from > 0) {
    from -= 1;
  }

  return {
    from,
    to,
    severity: "error",
    source: "Mermaid",
    message: getErrorMessage(error).slice(0, MESSAGE_LIMIT),
  };
};

export const mermaidLinter = linter(
  async (view) => {
    const state = view.state;
    const code = state.doc.toString();
    if (!code.trim()) {
      return [];
    }

    try {
      await validateMermaidCode(code);
      return [];
    } catch (error) {
      return [toDiagnostic(state, error)];
    }
  },
  { delay: 700 }
);

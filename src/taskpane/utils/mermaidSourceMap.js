// Mermaid preprocesses input before reporting parser locations. Keep the
// UTF-16 source offsets through the same removals Mermaid 12 performs in
// `preprocessDiagram` (cleanupText -> frontmatter -> directives -> comments).
//
// Recheck these rules against `cleanupText`/`frontMatterRegex`/`directiveRegex`
// in mermaidAPI when upgrading the engine; a mismatch only shifts the underline
// of a parse error, so this stays a best-effort mapper.
export function getMermaidSourcePosition(code, line, column) {
  let text = code;
  let offsets = Array.from({ length: code.length }, (_, index) => index);

  const replace = (pattern, replacement) => {
    const parts = [];
    const mapped = [];
    let start = 0;
    // matchAll requires the global flag; every rule below has it.
    for (const match of text.matchAll(pattern)) {
      const value = typeof replacement === "function" ? replacement(match[0]) : replacement;
      parts.push(text.slice(start, match.index), value);
      mapped.push(
        offsets.slice(start, match.index),
        Array.from(
          { length: value.length },
          (_, index) => offsets[match.index + Math.min(index, match[0].length - 1)]
        )
      );
      start = match.index + match[0].length;
    }
    parts.push(text.slice(start));
    mapped.push(offsets.slice(start));
    text = parts.join("");
    offsets = mapped.flat();
  };

  // cleanupText: normalise line endings, then single-quote HTML attributes.
  replace(/\r\n?/g, "\n");
  replace(/<(\w+)([^>]*)>/g, (tag) => tag.replace(/="([^"]*)"/g, "='$1'"));
  // processFrontmatter: leading indentation is captured, so the fence can be
  // nested inside another document.
  replace(/^([^\S\n\r]*)-{3}\s*[\n\r](.*?)[\n\r]\1-{3}\s*[\n\r]+/gs, "");
  // processDirectives: every `%%{init: ...}%%` and `%%{wrap}%%`.
  replace(/%{2}{\s*(?:(\w+)\s*:|(\w+))\s*(?:(\w+)|((?:(?!}%{2}).|\r?\n)*))?\s*(?:}%{2})?/gi, "");
  // cleanupComments: whole-line comments, then trim the leading whitespace.
  replace(/^\s*%%(?!{)[^\n]+\n?/gm, "");
  replace(/^\s+/, "");

  const lines = text.split("\n");
  const lineIndex = Math.min(Math.max(line - 1, 0), lines.length - 1);
  let position = Math.min(Math.max(column, 0), lines[lineIndex].length);
  for (let index = 0; index < lineIndex; index += 1) {
    position += lines[index].length + 1;
  }

  // Mermaid appends a newline before parsing, so an EOF location can lie just
  // beyond the mapped text. Keep it next to the last retained source token.
  return offsets[position] ?? (offsets.length ? offsets[offsets.length - 1] + 1 : 0);
}

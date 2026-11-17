// Mermaid 10.9 preprocesses input before reporting parser locations. Keep the
// UTF-16 source offsets through the same removals and entity substitutions.
// Recheck these rules against Mermaid's preprocessDiagram/encodeEntities when
// upgrading the engine. This mapping is only built after a parse failure.
export function getMermaidSourcePosition(code, line, column) {
  let text = code;
  let offsets = Array.from({ length: code.length }, (_, index) => index);

  const replace = (pattern, replacement) => {
    const parts = [];
    const mapped = [];
    let start = 0;
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

  replace(/\r\n?/g, "\n");
  replace(/<(\w+)([^>]*)>/g, (tag) => tag.replace(/="([^"]*)"/g, "='$1'"));
  replace(/^-{3}\s*[\n\r](.*?)[\n\r]-{3}\s*[\n\r]+/gs, "");
  replace(/%{2}{\s*(?:(\w+)\s*:|(\w+))\s*(?:(\w+)|((?:(?!}%{2}).|\r?\n)*))?\s*(?:}%{2})?/gi, "");
  replace(/^\s*%%(?!{)[^\n]+\n?/gm, "");
  replace(/^\s+/g, "");
  replace(/style.*:\S*#.*;/g, (match) => match.slice(0, -1));
  replace(/classDef.*:\S*#.*;/g, (match) => match.slice(0, -1));
  replace(/#\w+;/g, (match) => {
    const entity = match.slice(1, -1);
    return `${/^\+?\d+$/.test(entity) ? "ﬂ°°" : "ﬂ°"}${entity}¶ß`;
  });

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

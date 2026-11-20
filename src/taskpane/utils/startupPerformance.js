/* global performance */

// Keep only the first occurrence so editing and theme changes do not replace
// the startup milestones. Inspect them in the WebView's Performance panel.
export function markStartup(stage) {
  if (typeof performance === "undefined" || !performance.mark) {
    return;
  }
  const name = `addin:${stage}`;
  if (!performance.getEntriesByName(name, "mark").length) {
    performance.mark(name);
  }
}

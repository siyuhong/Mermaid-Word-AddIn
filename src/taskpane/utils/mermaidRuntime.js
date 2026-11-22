let mermaidRendererPromise;

export function loadMermaidRenderer() {
  if (!mermaidRendererPromise) {
    mermaidRendererPromise = import(
      /* webpackChunkName: "mermaid-renderer" */ "./mermaidRenderer"
    ).catch((error) => {
      mermaidRendererPromise = undefined;
      throw error;
    });
  }
  return mermaidRendererPromise;
}

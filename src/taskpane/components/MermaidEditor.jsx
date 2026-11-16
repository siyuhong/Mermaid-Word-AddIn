/* global Office */
import * as React from "react";
import {
  Field,
  Text,
  Button,
  makeStyles,
  tokens,
  Dropdown,
  Option,
} from "@fluentui/react-components";
import CodeMirror from "@uiw/react-codemirror";
import enhancedMermaidLanguage, { mermaidCompletion } from "./enhancedMermaidLanguage";
import { autocompletion } from "@codemirror/autocomplete";
import { EditorView } from "@codemirror/view";
import { insertDiagram, getSelectedImageAltText } from "../taskpane";
import { renderMermaidSvg } from "../utils/mermaidRenderer";
import { readSvgGeometry, fitIntoBox } from "../utils/diagramUtils";

import { AddRegular, SubtractRegular, ArrowResetRegular } from "@fluentui/react-icons";

const DEFAULT_DIAGRAM = `flowchart LR
A[Hard] -->|Text| B(Round)
B --> C{Decision}
C -->|One| D[Result 1]
C -->|Two| E[Result 2]
`;

// Mermaid 10 wraps requirement body text at a fixed 30 characters, but draws
// it into a rect whose default width is only 200px. Widen the rect so wrapped
// lines (about 210px at 14px) fit inside the box.
const REQUIREMENT_MIN_RECT_WIDTH = 300;

const MIN_ZOOM = 0.1;
const MAX_ZOOM = 10;
const ZOOM_STEP = 1.2;
const WHEEL_ZOOM_SENSITIVITY = 0.0015;

const clampZoom = (value) => Math.min(Math.max(value, MIN_ZOOM), MAX_ZOOM);

const getWheelDelta = (event, stage) => {
  const modeMultiplier =
    event.deltaMode === 1 ? 16 : event.deltaMode === 2 ? stage.clientHeight : 1;
  return event.deltaY * modeMultiplier;
};

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "24px 20px 32px",
    minHeight: "100vh",
    paddingBottom: "100px", // Extra padding to ensure content doesn't hide behind sticky buttons
  },
  description: {
    color: tokens.colorNeutralForeground2,
  },
  editorPreview: {
    display: "flex",
    flexDirection: "column",
    gap: "20px",
  },
  editorField: {
    width: "100%",
    boxSizing: "border-box",
  },
  codeMirrorWrapper: {
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    overflow: "hidden",
    boxSizing: "border-box",
  },
  previewHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "12px",
    gap: "12px",
  },
  themeSelector: {
    minWidth: "150px",
  },
  previewContainer: {
    width: "100%",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    minHeight: "240px",
    maxHeight: "400px",
    overflow: "hidden", // Hide scrollbars to manage pan/zoom manually
    boxSizing: "border-box",
    display: "flex", // Use flexbox for layout
    flexDirection: "column", // Stack children vertically
    userSelect: "none",
  },
  zoomControlsWrapper: {
    display: "flex",
    justifyContent: "flex-end",
    padding: "8px 16px 0 16px", // Padding at the top of the container
    zIndex: 1,
    backgroundColor: tokens.colorNeutralBackground2, // Ensure background is consistent
  },
  previewDiagramWrapper: {
    flexGrow: 1, // Allow this wrapper to take up remaining space
    overflow: "hidden", // Keep inner content from scrolling the whole container
    padding: "0 16px 16px 16px", // Apply padding here for the diagram content
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "grab",
  },
  zoomControls: {
    display: "flex",
    gap: "4px", // Reduced gap for smaller buttons
    alignItems: "center",
  },
  zoomButton: {
    minWidth: "28px", // Smaller button size
    width: "28px",
    height: "28px",
    padding: 0,
    justifyContent: "center",
  },
  previewContent: {
    width: "100%",
    maxWidth: "100%",
    overflow: "visible",
    boxSizing: "border-box",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    "& svg": {
      display: "block",
      maxWidth: "100%",
      height: "auto",
    },
  },
  error: {
    color: tokens.colorPaletteRedForeground1,
    fontWeight: tokens.fontWeightSemibold,
  },
  loading: {
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    height: "200px",
    color: tokens.colorNeutralForeground2,
  },
  buttonContainer: {
    position: "sticky",
    bottom: "0",
    backgroundColor: tokens.colorNeutralBackground1,
    padding: "16px 0",
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
    zIndex: 10,
    display: "flex",
    justifyContent: "flex-end",
    gap: "8px",
    marginTop: "16px",
  },
  zoomValue: {
    minWidth: "44px",
    textAlign: "center",
    color: tokens.colorNeutralForeground2,
  },
});

const getErrorMessage = (error, fallback) => {
  if (!error) {
    return fallback;
  }

  if (typeof error === "string") {
    return error;
  }

  if (error?.str) {
    return error.str;
  }

  if (error?.message) {
    return error.message;
  }

  return fallback;
};

const MermaidEditor = () => {
  const styles = useStyles();
  const [code, setCode] = React.useState(DEFAULT_DIAGRAM);
  const [error, setError] = React.useState("");
  const [renderedDiagram, setRenderedDiagram] = React.useState(null);
  const [isInserting, setIsInserting] = React.useState(false);
  const [theme, setTheme] = React.useState("default");
  const [hasSelectedImage, setHasSelectedImage] = React.useState(false);
  const [selectedImageCode, setSelectedImageCode] = React.useState("");
  const [isDragging, setIsDragging] = React.useState(false);
  const [startDragX, setStartDragX] = React.useState(0);
  const [startDragY, setStartDragY] = React.useState(0);
  const previewRef = React.useRef(null);
  const previewStageRef = React.useRef(null);
  const previewContainerRef = React.useRef(null);
  const previewAnimationFrameRef = React.useRef(null);
  const zoomLabelRef = React.useRef(null);
  const viewRef = React.useRef({ scale: 1, translationX: 0, translationY: 0 });
  const renderIndexRef = React.useRef(0);
  const renderVersionRef = React.useRef(0);
  const insertionInProgressRef = React.useRef(false);
  const debounceTimerRef = React.useRef(null);

  const canInsert =
    renderedDiagram?.code === code && renderedDiagram?.theme === theme && !isInserting;

  const invalidateRenderedDiagram = React.useCallback(() => {
    renderVersionRef.current += 1;
    setRenderedDiagram(null);
  }, []);

  const clearPreviewAnimationFrame = React.useCallback(() => {
    if (previewAnimationFrameRef.current) {
      cancelAnimationFrame(previewAnimationFrameRef.current);
      previewAnimationFrameRef.current = null;
    }
  }, []);

  const applyView = React.useCallback((nextView) => {
    viewRef.current = nextView;
    if (zoomLabelRef.current) {
      // Update the readout imperatively so wheel zooming never re-renders the
      // editor or CodeMirror.
      zoomLabelRef.current.textContent = `${Math.round(nextView.scale * 100)}%`;
    }
    if (previewRef.current) {
      previewRef.current.style.transform = `translate(${nextView.translationX}px, ${nextView.translationY}px) scale(${nextView.scale})`;
      previewRef.current.style.transformOrigin = "0 0";
    }
  }, []);

  const handleMouseDown = React.useCallback((e) => {
    if (e.button === 0) {
      // Left click
      setIsDragging(true);
      setStartDragX(e.clientX - viewRef.current.translationX);
      setStartDragY(e.clientY - viewRef.current.translationY);
      e.preventDefault();
    }
  }, []);

  const handleMouseMove = React.useCallback(
    (e) => {
      if (isDragging) {
        const { scale: currentScale } = viewRef.current;
        applyView({
          scale: currentScale,
          translationX: e.clientX - startDragX,
          translationY: e.clientY - startDragY,
        });
      }
    },
    [applyView, isDragging, startDragX, startDragY]
  );

  const handleMouseUp = React.useCallback(() => {
    setIsDragging(false);
  }, []);

  const handleMouseLeave = React.useCallback(() => {
    setIsDragging(false);
  }, []);

  const resetView = React.useCallback(() => {
    applyView({ scale: 1, translationX: 0, translationY: 0 });
  }, [applyView]);

  const zoomIn = React.useCallback(() => {
    const current = viewRef.current;
    applyView({ ...current, scale: clampZoom(current.scale * ZOOM_STEP) });
  }, [applyView]);

  const zoomOut = React.useCallback(() => {
    const current = viewRef.current;
    applyView({ ...current, scale: clampZoom(current.scale / ZOOM_STEP) });
  }, [applyView]);

  // Zoom around the pointer so the diagram stays anchored under the cursor.
  // A native non-passive listener is required because React's wheel listener
  // cannot prevent the task pane from scrolling.
  const handleWheel = React.useCallback(
    (event) => {
      const stage = previewStageRef.current;
      const preview = previewRef.current;
      if (!stage || !preview || !preview.querySelector("svg")) {
        return;
      }

      const delta = getWheelDelta(event, stage);
      if (!delta) {
        return;
      }

      event.preventDefault();
      setIsDragging(false);

      const current = viewRef.current;
      const nextScale = clampZoom(current.scale * Math.exp(-delta * WHEEL_ZOOM_SENSITIVITY));
      if (nextScale === current.scale) {
        return;
      }

      // getBoundingClientRect() already includes the current translation and
      // scale, so this converts the pointer into the preview's untransformed
      // coordinate space.
      const rect = preview.getBoundingClientRect();
      const anchorX = (event.clientX - rect.left) / current.scale;
      const anchorY = (event.clientY - rect.top) / current.scale;

      applyView({
        scale: nextScale,
        translationX: current.translationX - anchorX * (nextScale - current.scale),
        translationY: current.translationY - anchorY * (nextScale - current.scale),
      });
    },
    [applyView]
  );

  React.useEffect(() => {
    const stage = previewStageRef.current;
    if (!stage) {
      return undefined;
    }

    stage.addEventListener("wheel", handleWheel, { passive: false });
    return () => {
      stage.removeEventListener("wheel", handleWheel);
    };
  }, [handleWheel]);

  const resetPreviewSizing = React.useCallback(() => {
    clearPreviewAnimationFrame();
    applyView({ scale: 1, translationX: 0, translationY: 0 });

    if (previewRef.current) {
      previewRef.current.style.width = "100%";
      previewRef.current.style.maxWidth = "100%";
      previewRef.current.style.minWidth = "auto";
      previewRef.current.style.display = "";
      previewRef.current.style.justifyContent = "";
      previewRef.current.style.alignItems = "";
      previewRef.current.style.transform = ""; // Clear any existing transform
      previewRef.current.style.transformOrigin = "";
    }

    if (previewContainerRef.current) {
      previewContainerRef.current.scrollTop = 0;
      previewContainerRef.current.scrollLeft = 0;
    }
  }, [applyView, clearPreviewAnimationFrame]);

  const applyPreviewSizing = React.useCallback(() => {
    if (!previewRef.current || !previewStageRef.current) {
      return;
    }

    const svgElement = previewRef.current.querySelector("svg");
    if (!svgElement) {
      resetPreviewSizing();
      return;
    }

    // The SVG already carries explicit pixel dimensions from normalizeSvgSize.
    // Read those once and fit them into the real preview stage. No diagram-type
    // widths and no Word page dimensions are involved here.
    const geometry = readSvgGeometry(svgElement);
    const stage = previewStageRef.current;
    const rect = stage.getBoundingClientRect();
    const computedStyle = window.getComputedStyle(stage);
    const horizontalPadding =
      (parseFloat(computedStyle.paddingLeft) || 0) + (parseFloat(computedStyle.paddingRight) || 0);
    const verticalPadding =
      (parseFloat(computedStyle.paddingTop) || 0) + (parseFloat(computedStyle.paddingBottom) || 0);
    const availableWidth = Math.max(rect.width - horizontalPadding, 1);
    const availableHeight = Math.max(rect.height - verticalPadding, 1);
    // Base scale is the diagram's intrinsic size (1 SVG unit = 1 CSS px), so
    // the percentage readout means the same thing for preview and insertion.
    // Only shrink automatically when the intrinsic size does not fit.
    const fitted = fitIntoBox(
      geometry,
      { width: availableWidth, height: availableHeight },
      { allowUpscale: false }
    );

    previewRef.current.style.width = `${Math.max(1, Math.round(fitted.width))}px`;
    previewRef.current.style.height = `${Math.max(1, Math.round(fitted.height))}px`;
    previewRef.current.style.maxWidth = "100%";
    previewRef.current.style.minWidth = "auto";

    // Keep the intrinsic viewBox; the SVG simply fills the fitted box.
    svgElement.style.width = "100%";
    svgElement.style.height = "100%";
  }, [resetPreviewSizing]);

  const schedulePreviewSizing = React.useCallback(() => {
    clearPreviewAnimationFrame();
    previewAnimationFrameRef.current = requestAnimationFrame(() => {
      applyPreviewSizing();
    });
  }, [applyPreviewSizing, clearPreviewAnimationFrame]);

  // Only content/styling options here. Sizing is decided once by the shared
  // measurement viewport and fitIntoBox, never per diagram type.
  const mermaidConfig = React.useMemo(
    () => ({
      startOnLoad: false,
      securityLevel: "strict",
      suppressErrorRendering: true,
      theme: theme,
      fontFamily: "Arial, sans-serif",
      fontSize: 16,
      flowchart: {
        htmlLabels: true,
        curve: "basis",
      },
      sequence: {
        wrap: true,
        boxMargin: 10,
        boxTextMargin: 5,
        noteMargin: 10,
        messageMargin: 35,
      },
      gantt: {
        titleTopMargin: 25,
        barHeight: 20,
        fontSize: 12,
        sectionFontSize: 12,
        gridLineStartPadding: 35,
        gridLineEndPadding: 35,
        barGap: 4,
        topPadding: 30,
        leftPadding: 50,
        rightPadding: 50,
        bottomPadding: 30,
      },
      pie: {
        textPosition: 0.75,
        pieStrokeWidth: 2,
        pieOuterStrokeWidth: 2,
        pieInnerStrokeWidth: 2,
      },
      journey: {
        boxMargin: 10,
        boxTextMargin: 5,
      },
      // Mermaid 10 wraps requirement body text at 30 characters into a box
      // whose default width is 200px. This only fixes content overflow; it is
      // not part of the preview/insert sizing path.
      requirement: {
        rect_min_width: REQUIREMENT_MIN_RECT_WIDTH,
        rect_min_height: 200,
        rect_padding: 10,
        line_height: 20,
      },
    }),
    [theme]
  );

  React.useEffect(() => {
    let isActive = true;
    const version = renderVersionRef.current;
    setRenderedDiagram(null);

    const renderDiagram = async () => {
      if (!previewRef.current) {
        return;
      }

      try {
        const renderId = `mermaid-diagram-${renderIndexRef.current++}`;
        console.log("Rendering Mermaid diagram with ID:", renderId);
        const { svg } = await renderMermaidSvg(renderId, code, mermaidConfig);
        console.log("Mermaid diagram rendered successfully, SVG length:", svg.length);

        if (!isActive || version !== renderVersionRef.current || !previewRef.current) {
          return;
        }

        previewRef.current.innerHTML = svg;
        setError("");
        setRenderedDiagram({ svg, code, theme, config: mermaidConfig, version });

        schedulePreviewSizing();
      } catch (err) {
        if (!isActive || version !== renderVersionRef.current) {
          return;
        }

        console.error("Mermaid render error:", err);
        if (previewRef.current) {
          previewRef.current.innerHTML = "";
        }
        resetPreviewSizing();

        setError(getErrorMessage(err, "Unable to render the Mermaid preview."));
        setRenderedDiagram(null);
      }
    };

    // Debounce rendering to improve performance
    if (debounceTimerRef.current) {
      clearTimeout(debounceTimerRef.current);
    }

    debounceTimerRef.current = setTimeout(() => {
      renderDiagram();
    }, 200);

    return () => {
      isActive = false;
      if (debounceTimerRef.current) {
        clearTimeout(debounceTimerRef.current);
      }
    };
  }, [code, theme, mermaidConfig, schedulePreviewSizing, resetPreviewSizing]);

  // Resize the existing SVG without re-parsing or replacing the diagram.
  React.useEffect(() => {
    window.addEventListener("resize", schedulePreviewSizing);
    return () => {
      window.removeEventListener("resize", schedulePreviewSizing);
    };
  }, [schedulePreviewSizing]);

  React.useEffect(
    () => () => {
      clearPreviewAnimationFrame();
    },
    [clearPreviewAnimationFrame]
  );

  // Check for selected images when component mounts or when selection might change
  React.useEffect(() => {
    let isMounted = true;
    let initialTimeoutId = null;
    let selectionHandlerRegistered = false;
    let selectionCheckInProgress = false;
    let selectionCheckPending = false;

    const checkSelectedImage = async () => {
      if (!isMounted) {
        return;
      }
      if (selectionCheckInProgress) {
        selectionCheckPending = true;
        return;
      }

      selectionCheckInProgress = true;
      selectionCheckPending = false;

      try {
        const mermaidCode = await getSelectedImageAltText();
        if (!isMounted || selectionCheckPending) {
          return;
        }

        if (mermaidCode) {
          const normalizedCode = mermaidCode.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
          setHasSelectedImage(true);
          setSelectedImageCode((prev) => (prev === normalizedCode ? prev : normalizedCode));
        } else {
          setHasSelectedImage(false);
          setSelectedImageCode((prev) => (prev ? "" : prev));
        }
      } catch (error) {
        if (!isMounted || selectionCheckPending) {
          return;
        }

        // If there's no selection or it's not a Mermaid image, reset state
        setHasSelectedImage(false);
        setSelectedImageCode((prev) => (prev ? "" : prev));
      } finally {
        selectionCheckInProgress = false;
        if (isMounted && selectionCheckPending) {
          checkSelectedImage();
        }
      }
    };

    const registerSelectionChangedHandler = () => {
      if (
        typeof Office === "undefined" ||
        !Office?.context?.document?.addHandlerAsync ||
        !Office?.EventType?.DocumentSelectionChanged
      ) {
        return;
      }

      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        checkSelectedImage,
        (asyncResult) => {
          if (asyncResult?.status === Office.AsyncResultStatus.Succeeded) {
            selectionHandlerRegistered = true;
            if (!isMounted) {
              removeSelectionChangedHandler();
            }
          } else if (asyncResult?.error) {
            console.log("Failed to register selection changed handler:", asyncResult.error);
          }
        }
      );
    };

    const removeSelectionChangedHandler = () => {
      if (
        !selectionHandlerRegistered ||
        typeof Office === "undefined" ||
        !Office?.context?.document?.removeHandlerAsync ||
        !Office?.EventType?.DocumentSelectionChanged
      ) {
        return;
      }

      Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, {
        handler: checkSelectedImage,
      });
    };

    // Check immediately on mount and register event handler once Office is ready
    initialTimeoutId = setTimeout(() => {
      checkSelectedImage();
      registerSelectionChangedHandler();
    }, 100);

    // Add focus listener to check for selection changes when user returns to the task pane
    window.addEventListener("focus", checkSelectedImage);

    return () => {
      isMounted = false;
      if (initialTimeoutId) {
        clearTimeout(initialTimeoutId);
      }

      window.removeEventListener("focus", checkSelectedImage);

      removeSelectionChangedHandler();
      selectionCheckPending = false;
    };
  }, []);

  const handleCodeChange = (newCode) => {
    if (newCode === code) {
      return;
    }
    invalidateRenderedDiagram();
    setCode(newCode);
  };

  const handleThemeChange = (event, data) => {
    const newTheme = data.optionValue || "default";
    if (newTheme !== theme) {
      invalidateRenderedDiagram();
      setTheme(newTheme);
    }
  };

  const handleEditDiagram = () => {
    if (selectedImageCode) {
      handleCodeChange(selectedImageCode);
      setError(""); // Clear any previous errors
    }
  };

  const renderDiagramForExport = React.useCallback(async (diagram) => {
    const renderId = `mermaid-insert-safe-${renderIndexRef.current++}`;
    const { svg } = await renderMermaidSvg(renderId, diagram.code, diagram.config, true);
    return svg;
  }, []);

  const shouldRetryWithSafeSvg = React.useCallback((err) => {
    const errorText = `${err?.message || ""} ${err?.stack || ""}`;
    return /tainted canvases|securityerror|todataurl|failed to load svg image/i.test(errorText);
  }, []);

  const handleInsertDiagram = async () => {
    if (
      !canInsert ||
      renderedDiagram.version !== renderVersionRef.current ||
      insertionInProgressRef.current
    ) {
      return;
    }

    const diagram = renderedDiagram;
    insertionInProgressRef.current = true;
    setIsInserting(true);
    const getErrorMessage = (err) => err?.message || "Failed to insert diagram into Word.";

    try {
      setError(""); // Clear any previous errors
      // Always rasterise a render made for export: the preview SVG can contain
      // <foreignObject> labels, which taint the canvas and break toDataURL().
      const exportSvg = await renderDiagramForExport(diagram);
      await insertDiagram(exportSvg, diagram.code, { scale: viewRef.current.scale });
    } catch (err) {
      if (shouldRetryWithSafeSvg(err)) {
        try {
          console.warn("Retrying diagram insertion with the preview SVG...");
          await insertDiagram(diagram.svg, diagram.code, { scale: viewRef.current.scale });
          setError("");
          return;
        } catch (retryErr) {
          console.error("Insert diagram retry error:", retryErr);
          setError(getErrorMessage(retryErr));
          return;
        }
      }

      console.error("Insert diagram error:", err);
      setError(getErrorMessage(err));
    } finally {
      insertionInProgressRef.current = false;
      setIsInserting(false);
    }
  };

  const previewContainerClassName = styles.previewContainer;

  return (
    <section className={styles.root} aria-label="Mermaid editor">
      <div className={styles.editorPreview}>
        <Field className={styles.editorField}>
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              marginBottom: "8px",
            }}
          >
            <Text size="300" weight="semibold">
              Mermaid Code
            </Text>
          </div>
          <div className={styles.codeMirrorWrapper}>
            <CodeMirror
              value={code}
              onChange={handleCodeChange}
              height="260px"
              extensions={[
                enhancedMermaidLanguage,
                autocompletion({
                  override: [mermaidCompletion],
                  activateOnTyping: true,
                  maxRenderedOptions: 10,
                }),
                EditorView.lineWrapping,
              ]}
              basicSetup={{
                lineNumbers: true,
                foldGutter: true,
                highlightActiveLineGutter: true,
                highlightActiveLine: true,
              }}
            />
          </div>
        </Field>
        <div>
          <div className={styles.previewHeader}>
            <Text size="300" weight="semibold">
              Preview
            </Text>
            <Dropdown
              className={styles.themeSelector}
              value={theme}
              onOptionSelect={handleThemeChange}
              aria-label="Mermaid theme"
            >
              <Option value="default">Default</Option>
              <Option value="dark">Dark</Option>
              <Option value="forest">Forest</Option>
              <Option value="neutral">Neutral</Option>
            </Dropdown>
          </div>
          <div ref={previewContainerRef} className={previewContainerClassName}>
            <div className={styles.zoomControlsWrapper}>
              <div className={styles.zoomControls}>
                {" "}
                {/* Apply zoomControls style here */}
                <Button
                  size="small"
                  className={styles.zoomButton}
                  onClick={zoomIn}
                  aria-label="Zoom In"
                >
                  <AddRegular />
                </Button>
                <Button
                  size="small"
                  className={styles.zoomButton}
                  onClick={zoomOut}
                  aria-label="Zoom Out"
                >
                  <SubtractRegular />
                </Button>
                <Button
                  size="small"
                  className={styles.zoomButton}
                  onClick={resetView}
                  aria-label="Reset View"
                >
                  <ArrowResetRegular />
                </Button>
                <Text ref={zoomLabelRef} size="200" className={styles.zoomValue}>
                  100%
                </Text>
              </div>
            </div>
            <div
              ref={previewStageRef}
              className={styles.previewDiagramWrapper}
              onMouseDown={handleMouseDown}
              onMouseMove={handleMouseMove}
              onMouseUp={handleMouseUp}
              onMouseLeave={handleMouseLeave}
              style={{ cursor: isDragging ? "grabbing" : "grab" }}
            >
              <div ref={previewRef} className={styles.previewContent} />
            </div>
          </div>
        </div>
      </div>
      <div className={styles.buttonContainer}>
        <Button appearance="secondary" onClick={handleEditDiagram} disabled={!hasSelectedImage}>
          Edit Selected
        </Button>
        <Button appearance="primary" onClick={handleInsertDiagram} disabled={!canInsert}>
          Insert | Replace
        </Button>
      </div>
      {error && (
        <div role="alert" className={styles.error}>
          {error}
        </div>
      )}
    </section>
  );
};

export default MermaidEditor;

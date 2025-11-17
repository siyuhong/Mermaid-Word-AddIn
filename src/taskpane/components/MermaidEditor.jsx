/* global Office */
import * as React from "react";
import mermaid from "mermaid";
import {
  Field,
  Text,
  Button,
  makeStyles,
  mergeClasses,
  tokens,
  Dropdown,
  Option,
  SplitButton,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
} from "@fluentui/react-components";
import CodeMirror from "@uiw/react-codemirror";
import enhancedMermaidLanguage, { mermaidCompletion } from "./enhancedMermaidLanguage";
import { autocompletion } from "@codemirror/autocomplete";
import { EditorView } from "@codemirror/view";
import { insertDiagram, getSelectedImageAltText, getDocumentPageDimensions } from "../taskpane";
import {
  detectDiagramType,
  isGanttDiagram,
  isStateDiagram,
  isClassDiagram,
} from "../utils/diagramUtils";

const DEFAULT_DIAGRAM = `flowchart LR
A[Hard] -->|Text| B(Round)
B --> C{Decision}
C -->|One| D[Result 1]
C -->|Two| E[Result 2]
`;

const MERMAID_THEMES = {
  default: "default",
  dark: "dark",
  forest: "forest",
  neutral: "neutral",
};

const CUSTOM_THEME = {
  primary: "#0078d4",
  secondary: "#50e6ff",
  background: "#ffffff",
  text: "#000000",
};

const DEFAULT_PAGE_WIDTH_POINTS = 612;
const DEFAULT_PAGE_HEIGHT_POINTS = 792;
const DEFAULT_MARGIN_POINTS = 72;
const MIN_PREVIEW_SIZE_POINTS = 144;
const POINTS_PER_INCH = 72;
const DEFAULT_SCREEN_DPI = 96;

const createDefaultPageDimensions = () => ({
  pageWidth: DEFAULT_PAGE_WIDTH_POINTS,
  pageHeight: DEFAULT_PAGE_HEIGHT_POINTS,
  marginLeft: DEFAULT_MARGIN_POINTS,
  marginRight: DEFAULT_MARGIN_POINTS,
  marginTop: DEFAULT_MARGIN_POINTS,
  marginBottom: DEFAULT_MARGIN_POINTS,
});

const DEFAULT_PAGE_DIMENSIONS = createDefaultPageDimensions();

const getPixelsPerPoint = () => {
  if (typeof window !== "undefined") {
    const devicePixelRatio =
      typeof window.devicePixelRatio === "number" &&
      Number.isFinite(window.devicePixelRatio) &&
      window.devicePixelRatio > 0
        ? window.devicePixelRatio
        : 1;

    return (DEFAULT_SCREEN_DPI * devicePixelRatio) / POINTS_PER_INCH;
  }

  return DEFAULT_SCREEN_DPI / POINTS_PER_INCH;
};

const DEFAULT_ASPECT_RATIO = 4 / 3;

const calculatePreviewDisplaySize = (
  originalWidth,
  originalHeight,
  pageDimensions = DEFAULT_PAGE_DIMENSIONS
) => {
  const hasValidWidth = Number.isFinite(originalWidth) && originalWidth > 0;
  const hasValidHeight = Number.isFinite(originalHeight) && originalHeight > 0;

  const aspectRatio =
    hasValidWidth && hasValidHeight && originalHeight !== 0
      ? originalWidth / originalHeight
      : DEFAULT_ASPECT_RATIO;

  const availableWidth = Math.max(
    pageDimensions.pageWidth - pageDimensions.marginLeft - pageDimensions.marginRight,
    MIN_PREVIEW_SIZE_POINTS
  );

  const availableHeight = Math.max(
    pageDimensions.pageHeight - pageDimensions.marginTop - pageDimensions.marginBottom,
    MIN_PREVIEW_SIZE_POINTS
  );

  // Use 90% of available space for better text readability
  const maxUsableWidth = availableWidth * 0.9;
  const maxUsableHeight = availableHeight * 0.9;

  const availableAspectRatio = maxUsableWidth / maxUsableHeight;

  let targetWidth;
  let targetHeight;

  if (aspectRatio > availableAspectRatio) {
    targetWidth = maxUsableWidth;
    targetHeight = targetWidth / aspectRatio;
  } else {
    targetHeight = maxUsableHeight;
    targetWidth = targetHeight * aspectRatio;
  }

  // Ensure minimum size for text readability (2.5 inches instead of 2 inches)
  const MIN_READABLE_SIZE_POINTS = 180; // 2.5 inches
  if (targetWidth < MIN_READABLE_SIZE_POINTS && targetHeight < MIN_READABLE_SIZE_POINTS) {
    if (aspectRatio >= 1) {
      targetWidth = MIN_READABLE_SIZE_POINTS;
      targetHeight = targetWidth / aspectRatio;
    } else {
      targetHeight = MIN_READABLE_SIZE_POINTS;
      targetWidth = targetHeight * aspectRatio;
    }
  }

  // Ensure we don't exceed the available space
  if (targetWidth > maxUsableWidth) {
    targetWidth = maxUsableWidth;
    targetHeight = targetWidth / aspectRatio;
  }

  if (targetHeight > maxUsableHeight) {
    targetHeight = maxUsableHeight;
    targetWidth = targetHeight * aspectRatio;
  }

  console.log("Preview calculated with page dimensions:", {
    pageDimensions,
    availableWidth,
    availableHeight,
    maxUsableWidth,
    maxUsableHeight,
    targetWidth: Math.round(targetWidth),
    targetHeight: Math.round(targetHeight),
  });

  return {
    width: Math.round(targetWidth),
    height: Math.round(targetHeight),
  };
};

const extractSvgDimensions = (svgElement) => {
  if (!svgElement) {
    return { width: 0, height: 0 };
  }

  try {
    const viewBox = svgElement.getAttribute("viewBox");
    if (viewBox) {
      const viewBoxValues = viewBox.split(/\s+/).map(parseFloat);
      if (viewBoxValues.length >= 4) {
        const width = viewBoxValues[2];
        const height = viewBoxValues[3];
        if (Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0) {
          return { width, height };
        }
      }
    }

    const widthAttr = svgElement.getAttribute("width");
    const heightAttr = svgElement.getAttribute("height");

    if (widthAttr && heightAttr) {
      const width = parseFloat(widthAttr);
      const height = parseFloat(heightAttr);
      if (Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0) {
        return { width, height };
      }
    }

    if (svgElement.firstElementChild) {
      const childWidthAttr = svgElement.firstElementChild.getAttribute("width");
      const childHeightAttr = svgElement.firstElementChild.getAttribute("height");
      if (childWidthAttr && childHeightAttr) {
        const width = parseFloat(childWidthAttr);
        const height = parseFloat(childHeightAttr);
        if (Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0) {
          return { width, height };
        }
      }
    }

    if (svgElement.getBBox) {
      const bbox = svgElement.getBBox();
      if (bbox && bbox.width && bbox.height) {
        return { width: bbox.width, height: bbox.height };
      }
    }
  } catch (error) {
    console.log("Failed to extract SVG dimensions for preview:", error);
  }

  return { width: 800, height: 600 };
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
    padding: "16px",
    minHeight: "240px",
    maxHeight: "400px",
    overflowX: "auto",
    overflowY: "auto",
    boxSizing: "border-box",
    display: "flex",
    alignItems: "center",
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
  previewContentGantt: {
    "& svg": {
      minWidth: "800px",
      height: "auto",
      maxWidth: "none",
    },
  },
  previewContentStateDiagram: {
    "& svg": {
      maxWidth: "380px",
      width: "100%",
      height: "auto",
    },
  },
  previewContentClassDiagram: {
    "& svg": {
      maxWidth: "600px",
      width: "100%",
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
  const [canInsert, setCanInsert] = React.useState(false);
  const [theme, setTheme] = React.useState("default");
  const [hasSelectedImage, setHasSelectedImage] = React.useState(false);
  const [selectedImageCode, setSelectedImageCode] = React.useState("");
  const [diagramType, setDiagramType] = React.useState("");
  const [insertFormat, setInsertFormat] = React.useState("png");
  const [pageDimensions, setPageDimensions] = React.useState(() => ({
    ...DEFAULT_PAGE_DIMENSIONS,
  }));
  const previewRef = React.useRef(null);
  const previewContainerRef = React.useRef(null);
  const previewAnimationFrameRef = React.useRef(null);
  const renderIndexRef = React.useRef(0);
  const debounceTimerRef = React.useRef(null);
  const selectionCheckInProgressRef = React.useRef(false);
  const resizeTimerRef = React.useRef(null);
  const pageDimensionsRef = React.useRef({ ...DEFAULT_PAGE_DIMENSIONS });
  const pageDimensionsRequestRef = React.useRef(false);
  const lastPageDimensionsUpdateRef = React.useRef(0);

  const clearPreviewAnimationFrame = React.useCallback(() => {
    if (previewAnimationFrameRef.current) {
      cancelAnimationFrame(previewAnimationFrameRef.current);
      previewAnimationFrameRef.current = null;
    }
  }, []);

  const fetchPageDimensions = React.useCallback(async () => {
    const now = Date.now();
    const timeSinceLastUpdate = now - lastPageDimensionsUpdateRef.current;

    if (pageDimensionsRequestRef.current) {
      return pageDimensionsRef.current;
    }

    if (timeSinceLastUpdate < 2000) {
      return pageDimensionsRef.current;
    }

    try {
      pageDimensionsRequestRef.current = true;
      console.log("Fetching page dimensions from Word document...");

      const dims = await getDocumentPageDimensions();

      if (dims && typeof dims === "object") {
        const ensurePositive = (value, minValue, fallback) =>
          typeof value === "number" && Number.isFinite(value) && value > minValue
            ? value
            : fallback;

        const ensureNonNegative = (value, fallback) =>
          typeof value === "number" && Number.isFinite(value) && value >= 0 ? value : fallback;

        const normalizedDims = {
          pageWidth: ensurePositive(dims.pageWidth, 0, DEFAULT_PAGE_WIDTH_POINTS),
          pageHeight: ensurePositive(dims.pageHeight, 0, DEFAULT_PAGE_HEIGHT_POINTS),
          marginLeft: ensureNonNegative(dims.marginLeft, DEFAULT_MARGIN_POINTS),
          marginRight: ensureNonNegative(dims.marginRight, DEFAULT_MARGIN_POINTS),
          marginTop: ensureNonNegative(dims.marginTop, DEFAULT_MARGIN_POINTS),
          marginBottom: ensureNonNegative(dims.marginBottom, DEFAULT_MARGIN_POINTS),
        };

        pageDimensionsRef.current = normalizedDims;
        lastPageDimensionsUpdateRef.current = now;

        setPageDimensions((prev) => {
          if (
            prev.pageWidth !== normalizedDims.pageWidth ||
            prev.pageHeight !== normalizedDims.pageHeight ||
            prev.marginLeft !== normalizedDims.marginLeft ||
            prev.marginRight !== normalizedDims.marginRight ||
            prev.marginTop !== normalizedDims.marginTop ||
            prev.marginBottom !== normalizedDims.marginBottom
          ) {
            console.log("Page dimensions updated:", normalizedDims);
            return normalizedDims;
          }
          return prev;
        });

        return normalizedDims;
      }
    } catch (error) {
      console.log("Failed to fetch page dimensions, using cached/default:", error);
    } finally {
      pageDimensionsRequestRef.current = false;
    }

    return pageDimensionsRef.current;
  }, [setPageDimensions]);

  const resetPreviewSizing = React.useCallback(() => {
    clearPreviewAnimationFrame();

    if (previewRef.current) {
      previewRef.current.style.width = "100%";
      previewRef.current.style.maxWidth = "100%";
      previewRef.current.style.minWidth = "auto";
      previewRef.current.style.display = "";
      previewRef.current.style.justifyContent = "";
      previewRef.current.style.alignItems = "";
    }

    if (previewContainerRef.current) {
      previewContainerRef.current.scrollTop = 0;
      previewContainerRef.current.scrollLeft = 0;
    }
  }, [clearPreviewAnimationFrame]);

  const applyPreviewSizing = React.useCallback(() => {
    if (!previewRef.current || !previewContainerRef.current) {
      return;
    }

    const svgElement = previewRef.current.querySelector("svg");
    if (!svgElement) {
      resetPreviewSizing();
      return;
    }

    svgElement.removeAttribute("width");
    svgElement.removeAttribute("height");

    const { width: svgWidth, height: svgHeight } = extractSvgDimensions(svgElement);
    if (
      !Number.isFinite(svgWidth) ||
      !Number.isFinite(svgHeight) ||
      svgWidth <= 0 ||
      svgHeight <= 0
    ) {
      resetPreviewSizing();
      return;
    }

    const currentPageDims = pageDimensionsRef.current;
    const optimalSizePoints = calculatePreviewDisplaySize(svgWidth, svgHeight, currentPageDims);

    const pixelsPerPoint = getPixelsPerPoint();
    let targetWidthPx = Math.max(1, Math.round(optimalSizePoints.width * pixelsPerPoint));
    let targetHeightPx = Math.max(1, Math.round(optimalSizePoints.height * pixelsPerPoint));

    const containerElement = previewContainerRef.current;
    const containerRect = containerElement.getBoundingClientRect();
    const containerWidth =
      containerRect.width || containerElement.clientWidth || containerElement.offsetWidth || 0;
    const containerHeight =
      containerRect.height || containerElement.clientHeight || containerElement.offsetHeight || 0;

    let horizontalPadding = 0;
    let verticalPadding = 0;

    if (typeof window !== "undefined" && containerElement) {
      const computedStyle = window.getComputedStyle(containerElement);
      horizontalPadding =
        (parseFloat(computedStyle.paddingLeft) || 0) +
        (parseFloat(computedStyle.paddingRight) || 0);
      verticalPadding =
        (parseFloat(computedStyle.paddingTop) || 0) +
        (parseFloat(computedStyle.paddingBottom) || 0);
    }

    const availableWidth = Math.max(containerWidth - horizontalPadding, 1);
    const availableHeight = Math.max(containerHeight - verticalPadding, 1);

    let scale = 1;

    if (targetWidthPx > availableWidth && availableWidth > 0) {
      scale = Math.min(scale, availableWidth / targetWidthPx);
    }

    if (targetHeightPx > availableHeight && availableHeight > 0) {
      scale = Math.min(scale, availableHeight / targetHeightPx);
    }

    if (!Number.isFinite(scale) || scale <= 0) {
      scale = 1;
    }

    if (scale < 1) {
      targetWidthPx = Math.max(1, Math.round(targetWidthPx * scale));
      targetHeightPx = Math.max(1, Math.round(targetHeightPx * scale));
    }

    svgElement.style.width = `${targetWidthPx}px`;
    svgElement.style.height = `${targetHeightPx}px`;
    svgElement.style.maxWidth = `${targetWidthPx}px`;
    svgElement.style.maxHeight = `${targetHeightPx}px`;

    previewRef.current.style.width = `${targetWidthPx}px`;
    previewRef.current.style.maxWidth = "100%";
    previewRef.current.style.minWidth = "auto";
  }, [resetPreviewSizing]);

  const schedulePreviewSizing = React.useCallback(() => {
    clearPreviewAnimationFrame();
    previewAnimationFrameRef.current = requestAnimationFrame(() => {
      applyPreviewSizing();
    });
  }, [applyPreviewSizing, clearPreviewAnimationFrame]);

  // Initialize Mermaid configuration once
  const mermaidConfig = React.useMemo(
    () => ({
      startOnLoad: false,
      securityLevel: "strict",
      suppressErrorRendering: true,
      theme: theme,
      fontFamily: "Arial, sans-serif",
      fontSize: 16, // Increased from 16 for better readability
      flowchart: {
        useMaxWidth: true,
        htmlLabels: true,
        curve: "basis",
        padding: 20,
      },
      sequence: {
        useMaxWidth: true,
        wrap: true,
        boxMargin: 10,
        boxTextMargin: 5,
        noteMargin: 10,
        messageMargin: 35,
      },
      gantt: {
      useMaxWidth: false,
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
      bottomPadding: 30
      },
      classDiagram: {
        useMaxWidth: true,
        padding: 20,
      },
      state: {
        useMaxWidth: true,
        padding: 20,
      },
      stateDiagram: {
        useMaxWidth: true,
        padding: 20,
      },
      stateDiagramV2: {
        useMaxWidth: true,
        padding: 15,
        maxWidth: 480,
      },
      pie: {
        useMaxWidth: true,
        textPosition: 0.75,
        pieStrokeWidth: 2,
        pieOuterStrokeWidth: 2,
        pieInnerStrokeWidth: 2,
      },
      journey: {
        useMaxWidth: true,
        padding: 20,
        boxMargin: 10,
        boxTextMargin: 5,
        leftMargin: 100,
        rightMargin: 100,
      },
      er: {
        useMaxWidth: true,
        padding: 20,
      },
      mindmap: {
        useMaxWidth: true,
        padding: 20,
      },
      timeline: {
        useMaxWidth: true,
        padding: 20,
      },
      gitGraph: {
        useMaxWidth: true,
      },
    }),
    [theme]
  );

  React.useEffect(() => {
    fetchPageDimensions();
  }, [fetchPageDimensions]);

  React.useEffect(() => {
    console.log("Initializing Mermaid with theme:", theme);
    try {
      mermaid.initialize(mermaidConfig);
      console.log("Mermaid initialized successfully");
    } catch (error) {
      console.error("Error initializing Mermaid:", error);
    }
  }, [theme, mermaidConfig]);

  React.useEffect(() => {
    let isActive = true;

    const renderDiagram = async () => {
      if (!previewRef.current) {
        return;
      }

      await fetchPageDimensions();

      try {
        // First validate the Mermaid syntax
        await mermaid.parse(code);
        console.log("Mermaid code parsed successfully");
      } catch (err) {
        if (!isActive) {
          return;
        }

        console.error("Mermaid parse error:", err);
        previewRef.current.innerHTML = "";
        resetPreviewSizing();
        setError(getErrorMessage(err, "The provided Mermaid code is invalid."));
        setCanInsert(false);
        setDiagramType("");
        return;
      }

      try {
        const renderId = `mermaid-diagram-${renderIndexRef.current++}`;
        console.log("Rendering Mermaid diagram with ID:", renderId);
        const { svg } = await mermaid.render(renderId, code);
        console.log("Mermaid diagram rendered successfully, SVG length:", svg.length);

        if (!isActive || !previewRef.current) {
          return;
        }

        previewRef.current.innerHTML = svg;
        setError("");
        setCanInsert(true);
        setDiagramType(detectDiagramType(code));

        schedulePreviewSizing();
      } catch (err) {
        if (!isActive) {
          return;
        }

        console.error("Mermaid render error:", err);
        if (previewRef.current) {
          previewRef.current.innerHTML = "";
        }
        resetPreviewSizing();

        setError(getErrorMessage(err, "Unable to render the Mermaid preview."));
        setCanInsert(false);
        setDiagramType("");
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
  }, [code, theme, schedulePreviewSizing, resetPreviewSizing, fetchPageDimensions]);

  // Handle window resize to re-render diagram with new dimensions
  React.useEffect(() => {
    const handleResize = async () => {
      if (resizeTimerRef.current) {
        clearTimeout(resizeTimerRef.current);
      }

      resizeTimerRef.current = setTimeout(async () => {
        // Trigger re-render when window size changes
        if (previewRef.current && code) {
          try {
            const renderId = `mermaid-diagram-${renderIndexRef.current++}`;
            const { svg } = await mermaid.render(renderId, code);
            previewRef.current.innerHTML = svg;
            schedulePreviewSizing();
          } catch (err) {
            console.log("Error re-rendering on resize:", err);
          }
        }
      }, 300);
    };

    window.addEventListener("resize", handleResize);

    return () => {
      window.removeEventListener("resize", handleResize);
      if (resizeTimerRef.current) {
        clearTimeout(resizeTimerRef.current);
      }
    };
  }, [code, schedulePreviewSizing]);

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
    let intervalId = null;
    let selectionHandlerRegistered = false;

    const checkSelectedImage = async () => {
      if (selectionCheckInProgressRef.current) {
        return;
      }

      selectionCheckInProgressRef.current = true;

      try {
        fetchPageDimensions().catch(() => {});

        const mermaidCode = await getSelectedImageAltText();
        if (!isMounted) {
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
        if (!isMounted) {
          return;
        }

        // If there's no selection or it's not a Mermaid image, reset state
        setHasSelectedImage(false);
        setSelectedImageCode((prev) => (prev ? "" : prev));
      } finally {
        selectionCheckInProgressRef.current = false;
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

      Office.context.document.removeHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        checkSelectedImage
      );
    };

    // Check immediately on mount and register event handler once Office is ready
    initialTimeoutId = setTimeout(() => {
      checkSelectedImage();
      registerSelectionChangedHandler();
    }, 100);

    // Regular polling as a fallback in case events are missed
    intervalId = setInterval(checkSelectedImage, 3000);

    return () => {
      isMounted = false;
      if (initialTimeoutId) {
        clearTimeout(initialTimeoutId);
      }
      if (intervalId) {
        clearInterval(intervalId);
      }
      removeSelectionChangedHandler();
      selectionCheckInProgressRef.current = false;
    };
  }, [fetchPageDimensions]);

  const handleCodeChange = (newCode) => {
    setCode(newCode);
    setDiagramType(detectDiagramType(newCode));
  };

  const handleThemeChange = (event, data) => {
    setTheme(data.optionValue || "default");
  };

  const handleFormatChange = (event, data) => {
    setInsertFormat(data.optionValue || "png");
  };

  const handleEditDiagram = () => {
    if (selectedImageCode) {
      setCode(selectedImageCode);
      setDiagramType(detectDiagramType(selectedImageCode));
      setError(""); // Clear any previous errors
    }
  };

  const handleInsertDiagram = async (format = null) => {
    if (!canInsert || !previewRef.current) {
      return;
    }

    try {
      setError(""); // Clear any previous errors
      const svgContent = previewRef.current.innerHTML;
      const selectedFormat = format !== null ? format : insertFormat;
      await insertDiagram(svgContent, code, selectedFormat);
    } catch (err) {
      console.error("Insert diagram error:", err);
      const errorMessage = err?.message || "Failed to insert diagram into Word.";
      setError(errorMessage);
    }
  };

  const handleLoadExample = (example) => {
    setCode(example);
    setDiagramType(detectDiagramType(example));
    setError(""); // Clear any previous errors
  };

  const previewContainerClassName = styles.previewContainer;

  const previewContentClassName = mergeClasses(
    styles.previewContent,
    isGanttDiagram(diagramType) ? styles.previewContentGantt : undefined,
    isStateDiagram(diagramType) ? styles.previewContentStateDiagram : undefined,
    isClassDiagram(diagramType) ? styles.previewContentClassDiagram : undefined
  );

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
            <div style={{ display: "flex", gap: "8px" }}>
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
              <Dropdown
                className={styles.themeSelector}
                value={insertFormat}
                onOptionSelect={handleFormatChange}
                aria-label="Insert format"
              >
                <Option value="png">PNG</Option>
                <Option value="svg">SVG</Option>
              </Dropdown>
            </div>
          </div>
          <div ref={previewContainerRef} className={previewContainerClassName}>
            <div ref={previewRef} className={previewContentClassName} />
          </div>
        </div>
      </div>
      <div className={styles.buttonContainer}>
        <Button appearance="secondary" onClick={handleEditDiagram} disabled={!hasSelectedImage}>
          Edit Selected
        </Button>
        <Menu>
          <MenuTrigger>
            <SplitButton
              appearance="primary"
              disabled={!canInsert}
              onClick={handleInsertDiagram}
              menuButton={{ disabled: !canInsert }}
            >
              Insert to Word ({insertFormat.toUpperCase()})
            </SplitButton>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem
                onClick={async () => {
                  await handleInsertDiagram("png");
                }}
              >
                Insert as PNG
              </MenuItem>
              <MenuItem
                onClick={async () => {
                  await handleInsertDiagram("svg");
                }}
              >
                Insert as SVG
              </MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>
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

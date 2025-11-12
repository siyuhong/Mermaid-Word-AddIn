/* global Office */
import * as React from "react";
import mermaid from "mermaid";
import { Field, Text, Button, makeStyles, mergeClasses, tokens, Dropdown, Option } from "@fluentui/react-components";
import CodeMirror from "@uiw/react-codemirror";
import enhancedMermaidLanguage from "./enhancedMermaidLanguage";
import { insertDiagram, getSelectedImageAltText } from "../taskpane";
import { detectDiagramType, isGanttDiagram, isStateDiagram } from "../utils/diagramUtils";

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

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "24px 20px 32px",
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
  },
  codeMirrorWrapper: {
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    overflow: "hidden",
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
    overflowX: "hidden",
    overflowY: "auto",
  },
  previewContainerScrollable: {
    overflowX: "auto",
    overflowY: "hidden",
  },
  previewContent: {
    width: "100%",
    overflow: "visible",
    "& svg": {
      display: "block",
    },
  },
  previewContentResponsive: {
    "& svg": {
      width: "100%",
      maxWidth: "100%",
      height: "auto",
    },
  },
  previewContentGantt: {
    "& svg": {
      width: "auto",
      minWidth: "640px",
      maxWidth: "none",
      height: "auto",
    },
  },
  previewContentStateDiagram: {
    "& svg": {
      width: "100%",
      maxWidth: "380px",
      margin: "0 auto",
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
  const previewRef = React.useRef(null);
  const renderIndexRef = React.useRef(0);
  const debounceTimerRef = React.useRef(null);
  const selectionCheckInProgressRef = React.useRef(false);
  const resizeTimerRef = React.useRef(null);

  // Initialize Mermaid configuration once
  const mermaidConfig = {
    startOnLoad: false,
    securityLevel: "strict",
    suppressErrorRendering: true,
    theme: theme,
    fontFamily: "Arial, sans-serif",
    fontSize: 16,
    flowchart: {
      useMaxWidth: false,
      htmlLabels: true,
      curve: "basis",
      padding: 20
    },
    sequence: {
      useMaxWidth: false,
      wrap: true,
      boxMargin: 10,
      boxTextMargin: 5,
      noteMargin: 10,
      messageMargin: 35
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
      useMaxWidth: false,
      padding: 20
    },
    state: {
      useMaxWidth: false,
      padding: 20
    },
    stateDiagram: {
      useMaxWidth: false,
      padding: 20
    },
    stateDiagramV2: {
      useMaxWidth: true,
      padding: 15,
      maxWidth: 480
    },
    pie: {
      useMaxWidth: false,
      textPosition: 0.75,
      pieStrokeWidth: 2,
      pieOuterStrokeWidth: 2,
      pieInnerStrokeWidth: 2
    },
    journey: {
      useMaxWidth: false,
      padding: 20,
      boxMargin: 10,
      boxTextMargin: 5,
      leftMargin: 100,
      rightMargin: 100
    }
  };

  React.useEffect(() => {
    console.log("Initializing Mermaid with theme:", theme);
    try {
      mermaid.initialize(mermaidConfig);
      console.log("Mermaid initialized successfully");
    } catch (error) {
      console.error("Error initializing Mermaid:", error);
    }
  }, [theme]);

  React.useEffect(() => {
    let isActive = true;

    const renderDiagram = async () => {
      if (!previewRef.current) {
        return;
      }

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
      } catch (err) {
        if (!isActive) {
          return;
        }

        console.error("Mermaid render error:", err);
        if (previewRef.current) {
          previewRef.current.innerHTML = "";
        }

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
  }, [code, theme]);

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
  }, [code]);

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
  }, []);


  const handleCodeChange = (newCode) => {
    setCode(newCode);
    setDiagramType(detectDiagramType(newCode));
  };

  const handleThemeChange = (event, data) => {
    setTheme(data.optionValue || "default");
  };

  const handleEditDiagram = () => {
    if (selectedImageCode) {
      setCode(selectedImageCode);
      setDiagramType(detectDiagramType(selectedImageCode));
      setError(""); // Clear any previous errors
    }
  };

  const handleInsertDiagram = async () => {
    if (!canInsert || !previewRef.current) {
      return;
    }

    try {
      setError(""); // Clear any previous errors
      const svgContent = previewRef.current.innerHTML;
      await insertDiagram(svgContent, code);
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

  const previewContainerClassName = mergeClasses(
    styles.previewContainer,
    isGanttDiagram(diagramType) ? styles.previewContainerScrollable : undefined
  );

  const previewContentClassName = mergeClasses(
    styles.previewContent,
    isGanttDiagram(diagramType) ? styles.previewContentGantt : styles.previewContentResponsive,
    isStateDiagram(diagramType) ? styles.previewContentStateDiagram : undefined
  );

  return (
    <section className={styles.root} aria-label="Mermaid editor">
      <div className={styles.editorPreview}>
        <Field className={styles.editorField}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "8px" }}>
            <Text size="300" weight="semibold">Mermaid Code</Text>
          </div>
          <div className={styles.codeMirrorWrapper}>
            <CodeMirror
              value={code}
              onChange={handleCodeChange}
              height="260px"
              extensions={[enhancedMermaidLanguage]}
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
            <Text size="300" weight="semibold">Preview</Text>
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
          <div className={previewContainerClassName}>
            <div ref={previewRef} className={previewContentClassName} />
          </div>
        </div>
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", marginTop: "16px", gap: "8px" }}>
        <Button 
          appearance="secondary" 
          onClick={handleEditDiagram}
          disabled={!hasSelectedImage}
        >
          Edit Selected
        </Button>
        <Button 
          appearance="primary" 
          onClick={handleInsertDiagram}
          disabled={!canInsert}
        >
          Insert to Word
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

import * as React from "react";
import mermaid from "mermaid";
import { Field, Text, Button, makeStyles, tokens, Dropdown, Option } from "@fluentui/react-components";
import CodeMirror from "@uiw/react-codemirror";
import { mermaidLanguage } from "codemirror-lang-mermaid";
import { insertDiagram, getSelectedImageAltText } from "../taskpane";

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
    minHeight: "300px",
    overflow: "auto",
  },
  previewContent: {
    width: "100%",
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
  const previewRef = React.useRef(null);
  const renderIndexRef = React.useRef(0);
  const debounceTimerRef = React.useRef(null);

  React.useEffect(() => {
    console.log("Initializing Mermaid with theme:", theme);
    try {
      mermaid.initialize({
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
          fontSize: 14,
          sectionFontSize: 14,
          gridLineStartPadding: 35,
          gridLineEndPadding: 35,
          barGap: 4,
          topPadding: 50,
          leftPadding: 75,
          rightPadding: 75,
          bottomPadding: 50
        },
        classDiagram: {
          useMaxWidth: false,
          padding: 20
        },
        stateDiagram: {
          useMaxWidth: false,
          padding: 20
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
      });
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
      }
    };

    // Debounce rendering to improve performance
    if (debounceTimerRef.current) {
      clearTimeout(debounceTimerRef.current);
    }
    
    debounceTimerRef.current = setTimeout(() => {
      renderDiagram();
    }, 300);

    return () => {
      isActive = false;
      if (debounceTimerRef.current) {
        clearTimeout(debounceTimerRef.current);
      }
    };
  }, [code, theme]);

  // Check for selected images when component mounts or when selection might change
  React.useEffect(() => {
    const checkSelectedImage = async () => {
      try {
        const altText = await getSelectedImageAltText();
        if (altText && (altText.includes("flowchart") || altText.includes("graph") || 
            altText.includes("sequenceDiagram") || altText.includes("classDiagram") ||
            altText.includes("stateDiagram") || altText.includes("gantt") ||
            altText.includes("pie") || altText.includes("journey"))) {
          setHasSelectedImage(true);
          setSelectedImageCode(altText);
        } else {
          setHasSelectedImage(false);
          setSelectedImageCode("");
        }
      } catch (error) {
        // If there's no selection or it's not a Mermaid image, reset state
        setHasSelectedImage(false);
        setSelectedImageCode("");
      }
    };

    // Check immediately
    checkSelectedImage();

    // Set up a periodic check (every 2 seconds) to detect selection changes
    const intervalId = setInterval(checkSelectedImage, 2000);

    return () => {
      clearInterval(intervalId);
    };
  }, []);

  const handleCodeChange = (newCode) => {
    setCode(newCode);
  };

  const handleThemeChange = (event, data) => {
    setTheme(data.optionValue || "default");
  };

  const handleEditDiagram = () => {
    if (selectedImageCode) {
      setCode(selectedImageCode);
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

  return (
    <section className={styles.root} aria-label="Mermaid editor">
      <div className={styles.editorPreview}>
        <Field className={styles.editorField}>
          <Text size="300" weight="semibold">Mermaid Code</Text>
          <div className={styles.codeMirrorWrapper}>
            <CodeMirror
              value={code}
              onChange={handleCodeChange}
              height="260px"
              extensions={[mermaidLanguage]}
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
          <div className={styles.previewContainer}>
              <div ref={previewRef} className={styles.previewContent} />
          </div>
        </div>
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", marginTop: "16px", gap: "8px" }}>
        {hasSelectedImage && (
          <Button 
            appearance="secondary" 
            onClick={handleEditDiagram}
          >
            Edit Selected
          </Button>
        )}
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

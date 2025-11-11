import * as React from "react";
import mermaid from "mermaid";
import { Field, Textarea, Text, makeStyles, tokens } from "@fluentui/react-components";

const DEFAULT_DIAGRAM = `graph TD
    A[Start] --> B{Is it working?}
    B -- Yes --> C[Celebrate]
    B -- No --> D[Debug]
    D --> B
`;

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
    flexDirection: "row",
    flexWrap: "wrap",
    gap: "20px",
  },
  editorField: {
    flex: "1 1 320px",
  },
  textarea: {
    minHeight: "260px",
  },
  previewContainer: {
    flex: "1 1 320px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    padding: "16px",
    minHeight: "260px",
    overflow: "auto",
  },
  previewContent: {
    width: "100%",
  },
  error: {
    color: tokens.colorPaletteRedForeground1,
    fontWeight: tokens.fontWeightSemibold,
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

// List of unsupported diagram types that cause chunk loading errors
const UNSUPPORTED_DIAGRAMS = [
  "sequenceDiagram",
  "gantt",
  "classDiagram",
  "stateDiagram",
  "pie",
  "journey",
];

// Check if the diagram code uses an unsupported diagram type
const isUnsupportedDiagram = (code) => {
  const trimmedCode = code.trim();
  return UNSUPPORTED_DIAGRAMS.some((diagram) => trimmedCode.startsWith(diagram));
};

const MermaidEditor = () => {
  const styles = useStyles();
  const [code, setCode] = React.useState(DEFAULT_DIAGRAM);
  const [error, setError] = React.useState("");
  const previewRef = React.useRef(null);
  const renderIndexRef = React.useRef(0);

  React.useEffect(() => {
    mermaid.initialize({
      startOnLoad: false,
      securityLevel: "strict",
      suppressErrorRendering: true,
    });
  }, []);

  React.useEffect(() => {
    let isActive = true;

    const renderDiagram = async () => {
      if (!previewRef.current) {
        return;
      }

      if (isUnsupportedDiagram(code)) {
        if (!isActive) {
          return;
        }

        previewRef.current.innerHTML = "";
        const diagramType = code.trim().split(/[\s\n]/)[0];
        setError(
          `The diagram type "${diagramType}" is not supported. ` +
            "Supported types include: flowchart, graph, erd, gitGraph, mindmap, xychart, sankey, timeline."
        );
        return;
      }

      try {
        await mermaid.parse(code);
      } catch (err) {
        if (!isActive) {
          return;
        }

        previewRef.current.innerHTML = "";
        setError(getErrorMessage(err, "The provided Mermaid code is invalid."));
        return;
      }

      try {
        const renderId = `mermaid-diagram-${renderIndexRef.current++}`;
        const { svg } = await mermaid.render(renderId, code);

        if (!isActive || !previewRef.current) {
          return;
        }

        previewRef.current.innerHTML = svg;
        setError("");
      } catch (err) {
        if (!isActive) {
          return;
        }

        if (previewRef.current) {
          previewRef.current.innerHTML = "";
        }

        setError(getErrorMessage(err, "Unable to render the Mermaid preview."));
      }
    };

    renderDiagram();

    return () => {
      isActive = false;
    };
  }, [code]);

  const handleCodeChange = (event) => {
    setCode(event.target.value);
  };

  return (
    <section className={styles.root} aria-label="Mermaid editor">
      <div>
        <Text size="400" weight="semibold">
          Mermaid diagram editor
        </Text>
        <p className={styles.description}>
          Update the Mermaid definition on the left to instantly refresh the diagram preview on the right.
        </p>
      </div>
      <div className={styles.editorPreview}>
        <Field className={styles.editorField} label="Mermaid code">
          <Textarea value={code} onChange={handleCodeChange} appearance="outline" className={styles.textarea} />
        </Field>
        <div className={styles.previewContainer}>
          <div ref={previewRef} className={styles.previewContent} />
        </div>
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

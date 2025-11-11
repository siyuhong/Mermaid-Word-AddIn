import * as React from "react";
import mermaid from "mermaid";
import { Field, Textarea, Text, Button, makeStyles, tokens } from "@fluentui/react-components";
import { insertDiagram } from "../taskpane";

const DEFAULT_DIAGRAM = `sequenceDiagram
    participant Alice
    participant Bob
    Alice->>John: Hello John, how are you?
    loop Healthcheck
        John->>John: Fight against hypochondria
    end
    Note right of John: Rational thoughts <br/>prevail!
    John-->>Alice: Great!
    John->>Bob: How about you?
    Bob-->>John: Jolly good!

%% Try other diagram types:
%% gantt
%%     title A Gantt Diagram
%%     dateFormat  YYYY-MM-DD
%%     section Section
%%     A task           :a1, 2014-01-01, 30d
%%     Another task     :after a1  , 20d
%%
%% classDiagram
%%     Class01 <|-- AveryLongClass : Cool
%%     Class03 *-- Class04
%%     Class05 o-- Class06
%%     Class07 .. Class08
%%     Class09 --> C2 : Where am i?
%%
%% stateDiagram-v2
%%     [*] --> Still
%%     Still --> [*]
%%     Still --> Moving
%%     Moving --> Still
%%     Moving --> Crash
%%     Crash --> [*]
%%
%% pie
%%     title What is your favorite pet?
%%     "Dogs" : 386
%%     "Cats" : 85
%%     "Rats" : 15
%%
%% journey
%%     title My working day
%%     section Go to work
%%       Make tea: 5: Me
%%       Go upstairs: 3: Me
%%       Do work: 1: Me, Cat
%%     section Go home
%%       Go downstairs: 5: Me
%%       Sit down: 5: Me
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

const MermaidEditor = () => {
  const styles = useStyles();
  const [code, setCode] = React.useState(DEFAULT_DIAGRAM);
  const [error, setError] = React.useState("");
  const [canInsert, setCanInsert] = React.useState(false);
  const previewRef = React.useRef(null);
  const renderIndexRef = React.useRef(0);

  React.useEffect(() => {
    mermaid.initialize({
      startOnLoad: false,
      securityLevel: "strict",
      suppressErrorRendering: true,
      // Enable all diagram types including those that were previously disabled
      theme: "default",
      flowchart: {
        useMaxWidth: true,
        htmlLabels: true,
        curve: "basis"
      },
      sequence: {
        useMaxWidth: true,
        wrap: true
      },
      gantt: {
        useMaxWidth: true,
        titleTopMargin: 25,
        barHeight: 20,
        fontSize: 11,
        sectionFontSize: 11
      },
      classDiagram: {
        useMaxWidth: true
      },
      stateDiagram: {
        useMaxWidth: true
      },
      pie: {
        useMaxWidth: true
      },
      journey: {
        useMaxWidth: true
      }
    });
  }, []);

  React.useEffect(() => {
    let isActive = true;

    const renderDiagram = async () => {
      if (!previewRef.current) {
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
        setCanInsert(false);
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
        setCanInsert(true);
      } catch (err) {
        if (!isActive) {
          return;
        }

        if (previewRef.current) {
          previewRef.current.innerHTML = "";
        }

        setError(getErrorMessage(err, "Unable to render the Mermaid preview."));
        setCanInsert(false);
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

  const handleInsertDiagram = async () => {
    if (!canInsert || !previewRef.current) {
      return;
    }

    try {
      const svgContent = previewRef.current.innerHTML;
      await insertDiagram(svgContent, code);
    } catch (err) {
      setError("Failed to insert diagram into Word.");
    }
  };

  return (
    <section className={styles.root} aria-label="Mermaid editor">
      <div>
        <Text size="400" weight="semibold">
          Mermaid diagram editor
        </Text>
        <p className={styles.description}>
          Update the Mermaid definition on the left to instantly refresh the diagram preview on the right. All diagram types are now supported, including sequence diagrams, Gantt charts, class diagrams, state diagrams, pie charts, and journey diagrams.
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
      <div style={{ display: "flex", justifyContent: "flex-end", marginTop: "16px" }}>
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

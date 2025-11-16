import * as React from "react";
import MermaidEditor from "./MermaidEditor";
import { makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
  },
});

const App = () => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <MermaidEditor />
    </div>
  );
};

export default App;

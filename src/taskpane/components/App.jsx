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

const App = ({ officeStatus }) => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <MermaidEditor officeStatus={officeStatus} />
    </div>
  );
};

export default App;

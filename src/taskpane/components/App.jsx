import * as React from "react";
import PropTypes from "prop-types";
import Header from "./Header";
import HeroList from "./HeroList";
import MermaidEditor from "./MermaidEditor";
import { makeStyles } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.
  const listItems = [
    {
      icon: <Ribbon24Regular />,
      primaryText: "Compose and preview Mermaid diagrams in real time",
    },
    {
      icon: <LockOpen24Regular />,
      primaryText: "Experiment with flowcharts, sequence diagrams, and more",
    },
    {
      icon: <DesignIdeas24Regular />,
      primaryText: "Share polished diagrams directly from your document",
    },
  ];

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={title} message="Welcome" />
      <HeroList message="Build and preview Mermaid diagrams without leaving Word." items={listItems} />
      <MermaidEditor />
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;

import * as React from "react";
import PropTypes from "prop-types";
import Header from "./Header";
import HeroList from "./HeroList";
import TextInsertion from "./TextInsertion";
import MermaidEditor from "./MermaidEditor";
import { makeStyles, Tab, TabList } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular, DrawShape24Regular } from "@fluentui/react-icons";
import { insertText } from "../taskpane";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
  tabContent: {
    padding: "16px 0",
  },
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  const [selectedTab, setSelectedTab] = React.useState("welcome");
  
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.
  const listItems = [
    {
      icon: <Ribbon24Regular />,
      primaryText: "Achieve more with Office integration",
    },
    {
      icon: <LockOpen24Regular />,
      primaryText: "Unlock features and functionality",
    },
    {
      icon: <DesignIdeas24Regular />,
      primaryText: "Create and visualize like a pro",
    },
  ];

  const handleTabSelect = React.useCallback((event, data) => {
    setSelectedTab(data.value);
  }, []);

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={title} message="Welcome" />
      
      <TabList selectedValue={selectedTab} onTabSelect={handleTabSelect}>
        <Tab id="welcome" value="welcome" icon={<Ribbon24Regular />}>
          Welcome
        </Tab>
        <Tab id="text" value="text" icon={<LockOpen24Regular />}>
          Text Insertion
        </Tab>
        <Tab id="mermaid" value="mermaid" icon={<DrawShape24Regular />}>
          Mermaid Diagrams
        </Tab>
      </TabList>

      <div className={styles.tabContent}>
        {selectedTab === "welcome" && (
          <HeroList message="Discover what this add-in can do for you today!" items={listItems} />
        )}
        
        {selectedTab === "text" && (
          <TextInsertion insertText={insertText} />
        )}
        
        {selectedTab === "mermaid" && (
          <MermaidEditor />
        )}
      </div>
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;

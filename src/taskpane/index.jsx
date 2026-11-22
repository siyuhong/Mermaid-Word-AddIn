import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webDarkTheme, webLightTheme } from "@fluentui/react-components";
import { isOfficeDark, subscribeOfficeThemeChange } from "./utils/officeTheme";
import { waitForOfficeReady } from "./utils/officeRuntime";
import { markStartup } from "./utils/startupPerformance";

/* global document, module, require */

markStartup("app-entry");

const rootElement = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

const RootApp = ({ AppComponent }) => {
  const [isDark, setIsDark] = React.useState(isOfficeDark);
  const [officeStatus, setOfficeStatus] = React.useState("connecting");

  React.useLayoutEffect(() => {
    markStartup("app-visible");
  }, []);

  React.useEffect(() => {
    let active = true;
    const initializeOffice = async () => {
      try {
        const info = await waitForOfficeReady();
        markStartup("office-ready");
        if (active) {
          setOfficeStatus(info.host === "Word" ? "ready" : "unavailable");
        }
      } catch {
        if (active) {
          setOfficeStatus("error");
        }
      }
    };
    initializeOffice();
    return () => {
      active = false;
    };
  }, []);

  React.useEffect(() => {
    setIsDark(isOfficeDark());
    return subscribeOfficeThemeChange(() => setIsDark(isOfficeDark()));
  }, [officeStatus]);

  return (
    <FluentProvider theme={isDark ? webDarkTheme : webLightTheme}>
      <AppComponent officeStatus={officeStatus} />
    </FluentProvider>
  );
};

// Editing and preview do not need the Word host. Only document operations
// wait for onReady; a slow handshake must not keep the entire pane blank.
root?.render(<RootApp AppComponent={App} />);

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(<RootApp AppComponent={NextApp} />);
  });
}

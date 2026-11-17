import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webDarkTheme, webLightTheme } from "@fluentui/react-components";
import { isOfficeDark, subscribeOfficeThemeChange } from "./utils/officeTheme";

/* global document, Office, module, require */

const title = "Contoso Task Pane Add-in";

const rootElement = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

const RootApp = ({ title: appTitle }) => {
  const [isDark, setIsDark] = React.useState(isOfficeDark);

  React.useEffect(() => {
    setIsDark(isOfficeDark());
    return subscribeOfficeThemeChange(() => setIsDark(isOfficeDark()));
  }, []);

  return (
    <FluentProvider theme={isDark ? webDarkTheme : webLightTheme}>
      <App title={appTitle} />
    </FluentProvider>
  );
};

/* Render application after Office initializes */
Office.onReady(() => {
  root?.render(<RootApp title={title} />);
});

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}

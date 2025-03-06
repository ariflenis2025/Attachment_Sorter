import * as React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import AppDialog from "./components/AppDialog";

/* global document, Office, module, require, HTMLElement */

const title = "Contoso Task Pane Add-in";

const rootElement: HTMLElement | null = document.getElementById("Dialogcontainer");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <AppDialog/>
    </FluentProvider>
  );
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/AppDialog", () => {
    const NextApp = require("./components/AppDialog").default;
    root?.render(NextApp);
  });
}

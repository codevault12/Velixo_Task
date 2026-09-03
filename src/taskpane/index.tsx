import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { App } from "./App";
import "./App.css";

/* global Office, document */

Office.onReady(() => {
  const container = document.getElementById("root");
  if (!container) {
    return;
  }
  createRoot(container).render(
    <FluentProvider theme={webLightTheme}>
      <App />
    </FluentProvider>
  );
});

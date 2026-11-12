import { registerOfficeShortcuts } from "./shortcuts";
import React from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import "./styles.css";
createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);

void registerOfficeShortcuts().catch(() => {
  /* Pane shortcuts remain available. */
});

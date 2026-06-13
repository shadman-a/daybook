import React from "react";
import ReactDOM from "react-dom/client";
import { PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import App from "./App";
import { msalConfig } from "./auth/msalConfig";
import "./styles.css";

const rootElement = document.getElementById("root");
if (!rootElement) throw new Error("Daybook could not find its root element.");

const root = ReactDOM.createRoot(rootElement);
const msalInstance = new PublicClientApplication(msalConfig);

async function start() {
  try {
    await msalInstance.initialize();
    const redirectResponse = await msalInstance.handleRedirectPromise();
    const account = redirectResponse?.account ?? msalInstance.getAllAccounts()[0];
    if (account) msalInstance.setActiveAccount(account);

    root.render(
      <React.StrictMode>
        <MsalProvider instance={msalInstance}>
          <App />
        </MsalProvider>
      </React.StrictMode>
    );
  } catch (error) {
    const message = error instanceof Error ? error.message : "Microsoft authentication could not start.";
    root.render(
      <main className="entry-screen">
        <div className="entry-card setup-card">
          <span className="eyebrow">Startup error</span>
          <h1>Daybook could not start</h1>
          <p>{message}</p>
        </div>
      </main>
    );
  }
}

void start();

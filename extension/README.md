# Daybook Copilot Bridge

## Install locally

1. Open `chrome://extensions` in Chrome.
2. Turn on **Developer mode**.
3. Choose **Load unpacked**.
4. Select this `extension` directory.
5. Keep Microsoft 365 Copilot signed in at `https://m365.cloud.microsoft/chat`.

The extension stores synchronized Copilot records and generated briefs in
`chrome.storage.local` for 90 days. It does not receive or store the Graph
Explorer access token.

The extension only sends Daybook data to Microsoft when **Export to Copilot**
is pressed. The generated Markdown attachment may be stored in OneDrive for
Business under Microsoft's normal Copilot upload behavior.

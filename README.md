# Daybook Starter

Teams-first Microsoft Graph daily timeline with local Microsoft 365 Copilot
history synchronization and daily-brief export.

## Run

Requires Node.js 18 or newer.
Microsoft Teams chat APIs require a Microsoft 365 work or school account.

```bash
npm install
cp .env.example .env.local
# put your Entra app client id in .env.local
npm run dev
```

## Use a Graph Explorer token

If you cannot register an Entra application, copy the access token from Graph
Explorer and paste it into Daybook's sign-in screen. Daybook keeps the token in
`sessionStorage`, so it survives refreshes in the same browser session. It is
cleared on sign-out, expiration, or when the session ends.

The token must include `Chat.Read` for the Teams timeline. Calendar, mail, and
file context appears only when the token includes the corresponding delegated
permissions.

## Entra app registration

Create a SPA app registration with redirect URI:

```txt
http://localhost:5173
```

Add delegated permissions:

```txt
User.Read
Chat.Read
Calendars.Read
Mail.Read
Files.Read.All
```

Teams messages are pulled with:

```txt
GET /me/chats
GET /me/chats/{chat-id}/messages
```

Calendar/email/files are secondary timeline context.

Daybook does not require application permissions or admin-only chat APIs. If a
secondary permission is not granted, Teams remains available and the app shows
a source-specific warning.

## Copilot extension

Microsoft does not provide a delegated, non-admin API for Copilot interaction
history. Daybook therefore includes a local Chrome extension in `extension/`.

1. Open `chrome://extensions`.
2. Enable **Developer mode**.
3. Select **Load unpacked** and choose the repository's `extension` directory.
4. Sign into Microsoft 365 Copilot at `https://m365.cloud.microsoft/chat`.
5. Reload Daybook, then use **Sync Copilot** or **Export to Copilot**.

The extension stores synchronized Copilot records and generated briefs in
Chrome local storage for 90 days. It never receives the Graph Explorer token.

**Export to Copilot** creates a Markdown snapshot before opening a new Copilot
chat. The export is marked and quarantined so the generated summary cannot be
ingested into a later export. Microsoft may retain uploaded attachments in
OneDrive for Business.

If Copilot's page layout or upload policy prevents automation, use **Download
Markdown** and attach the generated file manually.

Validate both parts with:

```bash
npm run build
npm run check:extension
```

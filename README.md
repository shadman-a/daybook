# Daybook Starter

Teams-first Microsoft Graph daily timeline.

## Run

Requires Node.js 18 or newer.
Microsoft Teams chat APIs require a Microsoft 365 work or school account.

```bash
npm install
cp .env.example .env.local
# put your Entra app client id in .env.local
npm run dev
```

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

Build out this Daybook starter app. Keep it Teams-first. Make the UI nicer, fix TypeScript issues, improve error states, and preserve the current architecture.

Core requirement: the app must sign in with Microsoft using MSAL, call `/me/chats`, then call `/me/chats/{chat-id}/messages` for active chats, filter messages to the selected day, and show them in a chronological timeline. Calendar, email, files, and Copilot link are secondary context.

Do not add a backend, database, admin Graph permissions, AI summaries, webhooks, or `/users/{id}/chats/getAllMessages` yet.

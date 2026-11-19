# Mail Assistant

A personal web app for reading Outlook (Microsoft 365) mail in a clean view and turning a short note into a polished reply with Claude.

**Live:** https://domantas16.github.io/mail-assistant/ (open **Demo** to try it with sample emails)

## What it does

- **Chats** (the start screen) or **Email**: switch at the top of the list. Chats puts every email with one person in one messenger-style conversation; pin important contacts to keep them at the top
- **A reply is ready when an email opens.** Claude writes it (or prepares it in the background), or, with AI off, the best-fitting standard answer is filled in. Customer emails waiting for a reply are marked *Draft ready*
- **Send or Schedule.** Send waits 5 seconds so it can be undone; Schedule sends tomorrow morning, Monday, or at a chosen time (Outlook's deferred delivery). Send warns if `[placeholders]` are left
- Change the draft with one line in any language ("shorter", "10 vnt 450 EUR, 3 sav."), or quick buttons
- **Shared mailboxes** (e.g. info@) you can already open in Outlook
- Inbox, folders and search; readable threads with quoted history hidden and remote images blocked until allowed
- Contacts: the people you email, plus saved Outlook contacts when allowed

## Permissions and privacy

The app starts **read-only** and asks only for what is switched on in Settings:

| Feature | Microsoft Graph permissions (delegated) |
| --- | --- |
| Always | `User.Read`, `Mail.Read`, `offline_access` |
| Send and schedule from the app | `Mail.ReadWrite`, `Mail.Send` |
| Shared mailboxes | `Mail.Read.Shared` (with sending: `Mail.ReadWrite.Shared`, `Mail.Send.Shared`) |
| Outlook contacts and people search | `Contacts.Read`, `People.Read` |

Without sending, **Reply in Outlook** copies the draft and opens the email in Outlook.

What Claude may see is a setting too:

- **Off:** nothing leaves Microsoft and the browser; drafts come from the standard answers
- **My draft only:** Claude sees your draft and instructions, never the customer's email
- **Full:** Claude reads the thread and writes the reply

## How it works

There is no server. The app is a set of static files hosted on GitHub Pages.

- **Microsoft 365:** OAuth 2.0 authorization code flow with PKCE for single-page apps; Microsoft Graph is called from the browser with delegated permissions, so the app can only reach what the signed-in user can.
- **Claude:** the Anthropic API is called with your own API key, only in the AI modes above.
- **Your data:** the API key, sign-in tokens, settings, standard answers and prepared drafts are stored only in your browser (localStorage). Nothing is stored in this repository.
- **Safe rendering:** email HTML is cleaned and shown in a sandboxed iframe with no scripts, under a strict Content Security Policy.

## Setup

1. **IT registers the app** in Microsoft Entra ID: single tenant, platform **Single-page application**, redirect URI `https://domantas16.github.io/mail-assistant/`, the delegated permissions above (start with the first row), no client secret.
2. Open the app, paste the **Application (client) ID** and **Directory (tenant) ID**, and sign in.
3. Optional: add a Claude API key from the [Claude Console](https://platform.claude.com/settings/keys) and pick the AI mode.
4. In **Settings**, adjust the standard answers, style guide and signature, and switch on any features IT approved.

## Files

| File | Purpose |
| --- | --- |
| `index.html`, `styles.css`, `icon.svg` | Page shell and design |
| `app.js` | Screens and interactions |
| `auth.js` | Microsoft sign-in (PKCE) |
| `graph.js` | Outlook access via Microsoft Graph |
| `ai.js` | Prompts and streaming drafts from Claude |
| `answers.js` | Standard answers and keyword matching |
| `demo.js` | Fictional sample mailbox for demo mode |
| `lib.js` | Helpers: storage, formatting, safe email rendering, icons |
| `config.js` | Optional built-in client and tenant IDs (not secret) |

To run it locally, serve the folder (for example `python -m http.server 8000`) and open http://localhost:8000/. Microsoft sign-in also needs `http://localhost:8000/` registered as a redirect URI.

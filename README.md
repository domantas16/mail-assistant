# Mail Assistant

A personal web app for reading Outlook (Microsoft 365) mail in a clean view and turning a short note into a polished reply with Claude.

**Live:** https://domantas16.github.io/mail-assistant/ (open **Demo** to try it with sample emails)

## What it does

- **Chats** (the start screen): every email with one person in one messenger-style conversation, newest at the bottom. Pin important contacts to keep them at the top
- **Shared mailboxes:** switch between your own mailbox and shared ones (e.g. info@) you can already open in Outlook
- Inbox, folders and search, grouped by conversation
- Readable threads: quoted history hidden, remote images blocked until you allow them
- Contacts: the people you email most, plus saved Outlook contacts, with recent emails per person
- **Reply with AI:** write a few words in any language, press *Draft reply*, then revise ("shorter", "more formal", or your own instruction)
- **Send** the reply from the app (with a confirm step), save it to Outlook Drafts, or copy it (reply or reply all)
- New emails drafted from a short note, with recipient suggestions

## How it works

There is no server. The app is a set of static files hosted on GitHub Pages.

- **Microsoft 365:** sign-in uses the OAuth 2.0 authorization code flow with PKCE for single-page apps. The app calls Microsoft Graph directly from the browser with delegated permissions, so it can only reach the signed-in user's own mailbox.
- **Sending** goes through Microsoft Graph as you (or as the shared mailbox), after you confirm. Sent mail lands in Sent Items as usual.
- **Claude:** when you press *Draft*, only that thread and your note go to the Anthropic API, using your own API key.
- **Your data:** the API key, sign-in tokens, style guide, notes and signature are stored only in your browser (localStorage). Nothing is stored in this repository.
- **Safe rendering:** email HTML is cleaned and shown in a sandboxed iframe with no scripts, under a strict Content Security Policy.

## Setup

1. **IT registers the app** in Microsoft Entra ID:
   - Supported account types: this organization only (single tenant)
   - Platform: **Single-page application**, redirect URI `https://domantas16.github.io/mail-assistant/`
   - Delegated Microsoft Graph permissions with admin consent: `User.Read`, `Mail.ReadWrite`, `Mail.Send`, `Mail.ReadWrite.Shared`, `Mail.Send.Shared`, `Contacts.Read`, `People.Read`, `offline_access`
   - Shared mailboxes work only where the user already has Full Access (and Send As, for sending) in Exchange
   - No client secret
2. Open the app and paste the **Application (client) ID** and **Directory (tenant) ID**, then sign in.
3. Add a Claude API key from the [Claude Console](https://platform.claude.com/settings/keys).
4. In **Settings**, fill in your role, style guide, private notes, Outlook signature and any shared mailboxes.

## Files

| File | Purpose |
| --- | --- |
| `index.html`, `styles.css`, `icon.svg` | Page shell and design |
| `app.js` | Screens and interactions |
| `auth.js` | Microsoft sign-in (PKCE) |
| `graph.js` | Outlook access via Microsoft Graph |
| `ai.js` | Prompts and streaming drafts from Claude |
| `demo.js` | Fictional sample mailbox for demo mode |
| `lib.js` | Helpers: storage, formatting, safe email rendering, icons |
| `config.js` | Optional built-in client and tenant IDs (not secret) |

To run it locally, serve the folder (for example `python -m http.server 8000`) and open http://localhost:8000/. Microsoft sign-in also needs `http://localhost:8000/` registered as a redirect URI.

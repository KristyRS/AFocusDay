# A Focus Day

A calm, all-in-one productivity app — tasks, habits, a calendar, mood
tracking, and a focus timer, with real Microsoft/Google sign-in and
Claude-powered inbox summaries that turn your emails into to-dos and
events. The Anthropic API key lives only on the server (`server.js`) — it
is never sent to the browser.

## Just want to use it?

- **Web**: [afocusday.onrender.com](https://afocusday.onrender.com)
- **Desktop (Windows)**: download the installer from
  [Releases](https://github.com/KristyRS/AFocusDay/releases) — it opens
  the hosted app above in its own window, so there's no setup or API key
  needed on your end. The installer isn't code-signed, so Windows
  SmartScreen will warn you the first time — click **"More info" → "Run
  anyway."**

Sign in with Microsoft or Google and your data follows your account
across devices. Everything below is only needed if you want to run your
own copy for development.

## Features

- **Tasks** — due dates, subtasks, categories, inline editing
- **Habits** — daily check-ins, streaks, weekly targets
- **Calendar** — day/month/year views, multi-day events, and an AI scan
  that proposes events/deadlines found in your inbox
- **Focus timer** — Pomodoro-style sessions with optional breaks and
  cycles
- **Mood tracking** — log how you're feeling anytime, see the trend over
  the day
- **Inbox** — real Outlook mail, Claude-generated summaries and to-dos
- **Stats** — focus time, completed/unfinished tasks, and habit history
  by day

## Running your own copy

### 1. Configure Microsoft sign-in

You need an Azure AD (Entra ID) app registration. Set it up like this in
the [Azure Portal](https://portal.azure.com) → **Entra ID → App
registrations → your app**:

- **Authentication** → add a platform → **Single-page application (SPA)**
  → redirect URI `https://afocusday.onrender.com` (must match
  `MSAL_REDIRECT_URI` below exactly).
- **API permissions** → add **Microsoft Graph → Delegated → `User.Read`**
  and **`Mail.Read`**. If your org requires admin consent, click
  **Grant admin consent** (or ask whoever manages the tenant to).
- Copy the **Application (client) ID** and **Directory (tenant) ID** from
  the app's **Overview** page.

### 2. Configure Google sign-in (optional)

Create an OAuth 2.0 Client ID (type: Web application) in the
[Google Cloud Console](https://console.cloud.google.com), with
`https://afocusday.onrender.com` as an authorized origin.

### 3. Configure the Anthropic API key

Get a key from [console.anthropic.com](https://console.anthropic.com) →
API Keys.

### 4. Fill in `.env`

```
cp .env.example .env
```

Then edit `.env`:

```
MSAL_CLIENT_ID=<your Application (client) ID>
MSAL_TENANT_ID=<your Directory (tenant) ID, or "common" for any Microsoft account>
MSAL_REDIRECT_URI=https://afocusday.onrender.com

GOOGLE_CLIENT_ID=<your Google OAuth Client ID, optional>

ANTHROPIC_API_KEY=<your Anthropic API key>
ANTHROPIC_MODEL=claude-sonnet-4-6

PORT=3000
```

`.env` is gitignored — it stays on your machine only.

### 5. Run it

```
npm install
npm start
```

Open https://afocusday.onrender.com. Sign in with Microsoft or Google, then go to
**Inbox** to see your actual recent mail and click **Generate summary &
to-dos** to have Claude read it and propose a summary and action items.

## Desktop app

The `electron/` folder wraps the hosted app in a native window — no
separate build of the backend, no API key shipped inside it. It points at
whatever URL is set as `APP_URL` in `electron/main.js` (currently
`https://afocusday.onrender.com`) — change that first if you're building
an installer for your own deployment. To build:

```
cd electron
npm install
npm run dist
```

The installer lands in `electron/dist/`.

## Notes

- If mail fetch fails (e.g. consent not granted yet, or a personal
  Microsoft account without the right scope approved), the Inbox screen
  shows the specific Graph error instead of silently falling back to fake
  data.
- Your tasks, habits, events, mood log, and settings are saved server-side
  per signed-in account (`/api/user-data`), so they follow you across
  devices and browser sessions — not just local to one browser.
- On Render's free tier, that saved data lives on an ephemeral disk and
  can be wiped on redeploys or after the service spins down from
  inactivity. Fine for testing; for real persistence you'd want a paid
  tier with a persistent disk, or a proper database.

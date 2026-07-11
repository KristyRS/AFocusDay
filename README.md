# A Focus Day

A small local web app: real Microsoft sign-in (MSAL) that reads your Outlook
inbox, and a real Claude API call that turns it into a summary + suggested
to-dos. The Anthropic API key lives only on the server (`server.js`) — it is
never sent to the browser.

## 1. Configure Microsoft sign-in

You said you already have an Azure AD (Entra ID) app registration. Make sure
it's set up like this in the [Azure Portal](https://portal.azure.com) →
**Entra ID → App registrations → your app**:

- **Authentication** → add a platform → **Single-page application (SPA)**
  → redirect URI `http://localhost:3000` (must match `MSAL_REDIRECT_URI`
  below exactly, including the port).
- **API permissions** → add **Microsoft Graph → Delegated → `User.Read`**
  and **`Mail.Read`**. If your org requires admin consent, click
  **Grant admin consent** (or ask whoever manages the tenant to).
- Copy the **Application (client) ID** and **Directory (tenant) ID** from
  the app's **Overview** page.

## 2. Configure the Anthropic API key

Get a key from [console.anthropic.com](https://console.anthropic.com) →
API Keys.

## 3. Fill in `.env`

```
cp .env.example .env
```

Then edit `.env`:

```
MSAL_CLIENT_ID=<your Application (client) ID>
MSAL_TENANT_ID=<your Directory (tenant) ID, or "common" for any Microsoft account>
MSAL_REDIRECT_URI=http://localhost:3000

ANTHROPIC_API_KEY=<your Anthropic API key>
ANTHROPIC_MODEL=claude-sonnet-4-6

PORT=3000
```

`.env` is gitignored — it stays on your machine only.

## 4. Run it

```
npm install
npm start
```

Open http://localhost:3000. Click **Sign in with Microsoft**, approve the
`User.Read` / `Mail.Read` consent prompt, and you'll land on the real Home
screen with your real name. Go to **Inbox** to see your actual recent
Outlook mail, then click **Generate summary & to-dos** to have Claude read
those emails and propose a summary and action items.

## Notes

- If mail fetch fails (e.g. consent not granted yet, or a personal
  Microsoft account without the right scope approved), the Inbox screen
  shows the specific Graph error instead of silently falling back to fake
  data.
- Tasks, the focus timer, stats, and settings are local-only (in-memory)
  and reset on page reload — only sign-in and the AI summary talk to real
  services.
- Everything else (tasks/timer/stats/settings) is unchanged from the
  original design, just ported from the throwaway mockup into a real
  Express + vanilla JS app so the login and AI pieces could be real.

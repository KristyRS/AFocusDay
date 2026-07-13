require('dotenv').config();
const express = require('express');
const fs = require('fs');
const path = require('path');
const Anthropic = require('@anthropic-ai/sdk');

const app = express();
const port = process.env.PORT || 3000;

const anthropic = process.env.ANTHROPIC_API_KEY
  ? new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY })
  : null;

app.use(express.json());
app.use(express.static('public'));

// Per-user data lives here as one JSON file per account (Microsoft or
// Google) — this is what survives a browser reset or a switch to a
// different device/browser, unlike localStorage which is tied to one
// browser profile. In production this should point at a mounted
// persistent disk (set DATA_DIR to its mount path) — the app directory
// itself gets replaced on every deploy, so anything written next to the
// code without a real disk behind it is lost on the next deploy.
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

// Never trust a client-supplied user id — ask the provider itself (Microsoft
// Graph or Google) who the token actually belongs to, so one signed-in user
// can't read or overwrite another user's saved data by just changing a
// request parameter. The client says which provider it's using; that's not
// a trust boundary since we still verify the token against the real
// provider — a lie here just means the verification call fails (401), not
// an impersonation.
async function verifyUserId(req) {
  const authHeader = req.headers.authorization || '';
  const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : null;
  if (!token) return null;
  const provider = req.headers['x-auth-provider'] === 'google' ? 'google' : 'microsoft';
  try {
    if (provider === 'google') {
      const meRes = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (!meRes.ok) return null;
      const me = await meRes.json();
      return typeof me.sub === 'string' ? `google_${me.sub}` : null;
    }
    const meRes = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!meRes.ok) return null;
    const me = await meRes.json();
    return typeof me.id === 'string' ? `ms_${me.id}` : null;
  } catch {
    return null;
  }
}

function userDataPath(userId) {
  if (!/^(ms|google)_[0-9a-zA-Z-]{5,60}$/.test(userId)) return null;
  return path.join(DATA_DIR, `${userId}.json`);
}

app.get('/api/user-data', async (req, res) => {
  const userId = await verifyUserId(req);
  if (!userId) return res.status(401).json({ error: 'Not authenticated.' });
  const filePath = userDataPath(userId);
  if (!filePath) return res.status(400).json({ error: 'Invalid user id.' });
  try {
    if (!fs.existsSync(filePath)) return res.json({ data: null });
    res.json({ data: JSON.parse(fs.readFileSync(filePath, 'utf8')) });
  } catch (err) {
    console.error('Failed to read user data:', err);
    res.status(500).json({ error: 'Failed to read saved data.' });
  }
});

app.post('/api/user-data', async (req, res) => {
  const userId = await verifyUserId(req);
  if (!userId) return res.status(401).json({ error: 'Not authenticated.' });
  const filePath = userDataPath(userId);
  if (!filePath) return res.status(400).json({ error: 'Invalid user id.' });
  const data = req.body && req.body.data;
  if (!data || typeof data !== 'object') return res.status(400).json({ error: 'Missing data.' });
  try {
    fs.writeFileSync(filePath, JSON.stringify(data));
    res.json({ ok: true });
  } catch (err) {
    console.error('Failed to save user data:', err);
    res.status(500).json({ error: 'Failed to save data.' });
  }
});

// Non-secret MSAL + Google config the browser needs to start a sign-in.
app.get('/api/config', (req, res) => {
  res.json({
    clientId: process.env.MSAL_CLIENT_ID || '',
    tenantId: process.env.MSAL_TENANT_ID || 'common',
    redirectUri: process.env.MSAL_REDIRECT_URI || `http://localhost:${port}`,
    googleClientId: process.env.GOOGLE_CLIENT_ID || '',
  });
});

// Real AI summary: the Anthropic API key never leaves the server.
app.post('/api/summarize', async (req, res) => {
  if (!anthropic) {
    return res.status(500).json({ error: 'ANTHROPIC_API_KEY is not configured on the server (.env).' });
  }

  const emails = Array.isArray(req.body.emails) ? req.body.emails : [];
  if (emails.length === 0) {
    return res.status(400).json({ error: 'No emails provided.' });
  }

  const emailBlock = emails
    .slice(0, 25)
    .map((m, i) => `${i + 1}. From: ${m.sender}\n   Subject: ${m.subject}\n   Snippet: ${m.snippet || ''}`)
    .join('\n\n');

  const now = new Date();
  const today = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;

  const prompt = `Today's date is ${today}. Here are recent emails from a user's inbox:\n\n${emailBlock}\n\nRespond with ONLY valid JSON (no markdown fences) in this exact shape:\n{"summary": "a concise 2-3 sentence plain-English summary of what needs attention", "todos": [{"name": "short actionable task", "source": "sender name", "date": "YYYY-MM-DD or null"}]}\nInclude at most 5 todos, only for emails that actually imply an action. If an email suggests or implies a due date or deadline (including relative phrases like "by Friday" or "in 2 weeks"), resolve it to an actual date using today's date as the reference and put it in "date" — otherwise use null.`;

  try {
    const model = process.env.ANTHROPIC_MODEL || 'claude-sonnet-4-6';
    const message = await anthropic.messages.create({
      model,
      max_tokens: 1024,
      messages: [{ role: 'user', content: prompt }],
    });

    const text = message.content.map((block) => (block.type === 'text' ? block.text : '')).join('');
    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch {
      const match = text.match(/\{[\s\S]*\}/);
      parsed = match ? JSON.parse(match[0]) : { summary: text, todos: [] };
    }

    const dateRe = /^\d{4}-\d{2}-\d{2}$/;
    const todos = (Array.isArray(parsed.todos) ? parsed.todos : []).map((t) => ({
      name: t.name,
      source: t.source,
      date: dateRe.test(t.date) ? t.date : null,
    }));

    res.json({
      summary: parsed.summary || '',
      todos,
    });
  } catch (err) {
    console.error('Anthropic API error:', err);
    res.status(502).json({ error: 'AI summary request failed: ' + err.message });
  }
});

// Real AI event extraction: finds genuine events/deadlines/appointments in recent emails.
app.post('/api/extract-events', async (req, res) => {
  if (!anthropic) {
    return res.status(500).json({ error: 'ANTHROPIC_API_KEY is not configured on the server (.env).' });
  }

  const emails = Array.isArray(req.body.emails) ? req.body.emails : [];
  if (emails.length === 0) {
    return res.status(400).json({ error: 'No emails provided.' });
  }

  const emailBlock = emails
    .slice(0, 50)
    .map((m, i) => `${i + 1}. From: ${m.sender}\n   Subject: ${m.subject}\n   Snippet: ${m.snippet || ''}`)
    .join('\n\n');

  const now = new Date();
  const today = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;

  const prompt = `Today's date is ${today}. Here are recent emails from a user's inbox:

${emailBlock}

Find emails that describe a genuine, specific event, meeting, appointment, or deadline the recipient is personally expected to attend or act on (e.g. a calendar invite, an interview, a bill due date, a school event). Resolve any relative dates ("tomorrow," "next Friday," "in 2 weeks") into an actual calendar date based on today's date. Ignore marketing/promotional emails and anything without a specific date.

Respond with ONLY valid JSON (no markdown fences) in this exact shape:
{"events": [{"name": "short event description", "date": "YYYY-MM-DD", "source": "sender name"}]}
Include at most 10 events. Only include an event if you can confidently resolve an actual date.`;

  try {
    const model = process.env.ANTHROPIC_MODEL || 'claude-sonnet-4-6';
    const message = await anthropic.messages.create({
      model,
      max_tokens: 1024,
      messages: [{ role: 'user', content: prompt }],
    });

    const text = message.content.map((block) => (block.type === 'text' ? block.text : '')).join('');
    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch {
      const match = text.match(/\{[\s\S]*\}/);
      parsed = match ? JSON.parse(match[0]) : { events: [] };
    }

    const dateRe = /^\d{4}-\d{2}-\d{2}$/;
    const events = (Array.isArray(parsed.events) ? parsed.events : [])
      .filter((e) => e && e.name && dateRe.test(e.date))
      .map((e) => ({ name: e.name, date: e.date, source: e.source || 'Unknown sender' }));

    res.json({ events });
  } catch (err) {
    console.error('Anthropic API error:', err);
    res.status(502).json({ error: 'Event extraction failed: ' + err.message });
  }
});

app.listen(port, () => {
  console.log(`A Focus Day running at http://localhost:${port}`);
  if (!process.env.MSAL_CLIENT_ID) {
    console.warn('Warning: MSAL_CLIENT_ID not set in .env — Microsoft sign-in will not work yet.');
  }
  if (!process.env.ANTHROPIC_API_KEY) {
    console.warn('Warning: ANTHROPIC_API_KEY not set in .env — AI summaries will not work yet.');
  }
});

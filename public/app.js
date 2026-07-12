/* A Focus Day — real Microsoft sign-in (MSAL) + real Claude-powered inbox summaries */

function esc(str) {
  return String(str == null ? '' : str).replace(/[&<>"']/g, (c) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[c]));
}

const AVATAR_COLORS = [
  'oklch(0.6 0.13 250)', 'oklch(0.6 0.13 20)', 'oklch(0.6 0.1 90)',
  'oklch(0.55 0.02 152)', 'oklch(0.6 0.12 320)', 'oklch(0.6 0.12 180)',
];
function colorForName(name) {
  let hash = 0;
  for (let i = 0; i < name.length; i++) hash = (hash * 31 + name.charCodeAt(i)) >>> 0;
  return AVATAR_COLORS[hash % AVATAR_COLORS.length];
}

function decodeGmailBase64Url(data) {
  try {
    const binary = atob(data.replace(/-/g, '+').replace(/_/g, '/'));
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return new TextDecoder('utf-8').decode(bytes);
  } catch {
    return '';
  }
}
// Strips <style>/<script> content (not just the tags — their contents are
// raw CSS/JS and are exactly the "random code" that leaked through before),
// HTML comments, then tags, then decodes the common entities so text like
// &nbsp;/&amp; doesn't show up literally.
function htmlToCleanText(html) {
  let text = html
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<!--[\s\S]*?-->/g, ' ')
    .replace(/<(br|\/p|\/div|\/tr|\/li)\s*\/?>/gi, '\n')
    .replace(/<[^>]+>/g, ' ');
  text = text
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&zwnj;|&zwj;|&#8203;/gi, '');
  return text
    .split('\n')
    .map((line) => line.replace(/[ \t]+/g, ' ').trim())
    .filter((line, idx, arr) => !(line === '' && arr[idx - 1] === ''))
    .join('\n')
    .trim();
}
// Microsoft Graph's plain-text conversion renders links as "[Label]<url>" —
// fine for a short legitimate link, but marketing email tracking links carry
// huge encoded query strings that read as a wall of "code" once inlined.
// The label is what a reader actually needs; the raw tracking URL is noise.
function cleanOutlookPlainText(text) {
  return text
    .replace(/\[([^\]\n]{0,120})\]\s*<[^>\n]+>/g, '$1')
    .replace(/<https?:\/\/[^>\n]+>/gi, '')
    .split('\n')
    .map((line) => line.replace(/[ \t]+/g, ' ').trim())
    .filter((line, idx, arr) => !(line === '' && arr[idx - 1] === ''))
    .join('\n')
    .trim();
}
// Gmail messages are a MIME tree, not a single body field — walk it looking
// for a plain-text part first, falling back to a cleaned-up HTML-to-text
// conversion so we never hand raw HTML/CSS/JS to the page (same
// plain-text-only approach as Outlook).
function extractGmailBody(payload) {
  if (!payload) return '';
  if (payload.mimeType === 'text/plain' && payload.body && payload.body.data) {
    return decodeGmailBase64Url(payload.body.data);
  }
  if (Array.isArray(payload.parts)) {
    const plain = payload.parts.find((p) => p.mimeType === 'text/plain');
    if (plain && plain.body && plain.body.data) return decodeGmailBase64Url(plain.body.data);
    // Only recurse into genuinely nested multipart containers — recursing
    // into a plain leaf part here would skip the html-stripping branch
    // below via extractGmailBody's own generic body.data fallback.
    for (const part of payload.parts) {
      if (Array.isArray(part.parts)) {
        const nested = extractGmailBody(part);
        if (nested) return nested;
      }
    }
    const html = payload.parts.find((p) => p.mimeType === 'text/html');
    if (html && html.body && html.body.data) {
      return htmlToCleanText(decodeGmailBase64Url(html.body.data));
    }
  }
  if (payload.body && payload.body.data) return decodeGmailBase64Url(payload.body.data);
  return '';
}

// Warm pastel palette for task-category tags — Work stays neutral, the rest
// cycle through peach/mint/lavender/blue so tags read as distinct at a glance.
const CATEGORY_TAG_PALETTE = [
  { hue: 75, chroma: 0.012 }, // neutral (Work)
  { hue: 55, chroma: 0.06 }, // peach (Personal)
  { hue: 155, chroma: 0.07 }, // mint (Health)
  { hue: 300, chroma: 0.045 }, // lavender (Design)
  { hue: 250, chroma: 0.05 }, // soft blue (Learning)
  { hue: 20, chroma: 0.07 }, // coral
];
const CATEGORY_TAG_FIXED = { Work: 0, Personal: 1, Health: 2, Design: 3, Learning: 4 };
function categoryColor(tag, dark) {
  let index = CATEGORY_TAG_FIXED[tag];
  if (index === undefined) {
    let hash = 0;
    for (let i = 0; i < tag.length; i++) hash = (hash * 31 + tag.charCodeAt(i)) >>> 0;
    index = hash % CATEGORY_TAG_PALETTE.length;
  }
  const c = CATEGORY_TAG_PALETTE[index];
  return dark
    ? { bg: `oklch(0.3 ${c.chroma + 0.02} ${c.hue})`, text: `oklch(0.82 ${c.chroma + 0.05} ${c.hue})` }
    : { bg: `oklch(0.91 ${c.chroma} ${c.hue})`, text: `oklch(0.4 ${Math.max(c.chroma, 0.08)} ${c.hue})` };
}

function formatDuration(totalSeconds) {
  const total = Math.max(0, Math.round(totalSeconds));
  const m = Math.floor(total / 60);
  const s = total % 60;
  if (m === 0) return `${s} second${s === 1 ? '' : 's'}`;
  if (s === 0) return `${m} minute${m === 1 ? '' : 's'}`;
  return `${m}m ${s}s`;
}

function dayKey(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function parseLocalDate(dateStr) {
  const [y, m, d] = dateStr.split('-').map(Number);
  return new Date(y, m - 1, d);
}

function formatTaskDate(dateStr) {
  if (!dateStr) return '';
  if (dateStr === dayKey(new Date())) return 'Today';
  return parseLocalDate(dateStr).toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
}

const DEFAULT_CATEGORIES = ['Work', 'Personal', 'Design', 'Health', 'Learning'];

const MOOD_OPTIONS = [
  { value: 'awful', emoji: '😢', label: 'Awful' },
  { value: 'bad', emoji: '😕', label: 'Not great' },
  { value: 'okay', emoji: '😐', label: 'Okay' },
  { value: 'good', emoji: '🙂', label: 'Good' },
  { value: 'great', emoji: '😄', label: 'Great' },
];

const HABIT_ICONS = ['💧', '🏃', '📚', '🧘', '☀️', '🥗', '😴', '✍️', '🎯', '🎨'];

const DEFAULT_TASKS = [];

const App = {
  msalInstance: null,
  msalConfig: null,
  account: null,
  googleClientId: null,
  googleAccessToken: null,
  authProvider: null, // 'microsoft' | 'google' | null
  history: {}, // real per-day focus record, persisted in localStorage — { 'YYYY-MM-DD': { focusSeconds, sessionsCompleted } }
  taskHistory: {}, // completed-task log, persisted — { 'YYYY-MM-DD': [{ name, tag }] }
  moodLog: {}, // mood check-ins, persisted — { 'YYYY-MM-DD': [{ time, mood }] }
  emailNotes: {}, // per-email notes, persisted — { [graphMessageId]: 'note text' }
  habits: [], // habit definitions, persisted — [{ id, name, icon }]
  habitLog: {}, // habit check-ins, persisted — { 'YYYY-MM-DD': { [habitId]: true } }

  state: {
    screen: 'onboarding', // onboarding | home | setup | timer | checkin | breakdown-prompt | breakdown-edit | complete | stats | inbox | calendar | settings
    tasks: DEFAULT_TASKS.map((t) => ({ ...t, subtasks: [...t.subtasks] })),
    selectedTaskId: null,
    expandedSubtaskTaskId: null,
    editingTaskId: null,
    editingSubtask: null, // { taskId, subId }
    datePickerTaskId: null,
    datePickerSubtask: null, // { taskId, subId }
    categories: [...DEFAULT_CATEGORIES],
    newTaskDraftSubtasks: [],
    newSetupTaskDraftSubtasks: [],
    lastLoggedMood: null,

    // Session configuration, set on the Setup screen
    focusMinutes: 25,
    breakEnabled: false,
    breakMinutes: 5,
    cycles: 1,

    sessionsToday: 0,
    streak: 0,
    totalFocusSecondsToday: 0,
    timerTaskId: null,
    currentCycleIndex: 0,
    currentPhase: 'focus', // 'focus' | 'break'
    phaseDurationMin: 25,
    phaseElapsedSeconds: 0,
    lastSessionFocusSeconds: 0,
    lastSessionCycles: 0,
    checkinCompletedFully: true,
    statsViewMonth: { year: new Date().getFullYear(), month: new Date().getMonth() },
    statsSelectedDate: dayKey(new Date()),

    events: [], // calendar events — { id, name, date, time, important, source: 'manual' | 'ai' }
    calendarSelectedDate: dayKey(new Date()),
    calendarViewMode: 'day', // day | month | year
    calendarViewMonth: { year: new Date().getFullYear(), month: new Date().getMonth() },
    calendarViewYear: new Date().getFullYear(),
    eventScanStage: 'idle', // idle | scanning | ready | error
    eventScanError: '',
    proposedEvents: [], // AI-suggested events awaiting accept/dismiss — { id, name, date, source }

    reminders: [], // { id, date, text }
    editingReminderId: null,
    newHabitIcon: HABIT_ICONS[0],
    newHabitTarget: 7,
    showNewHabitForm: false,

    remainingSeconds: 0,
    isPaused: false,
    interval: null,
    theme: 'light',
    settingsToggles: { autoMode: false },

    userName: '',
    userInitial: '',
    signingIn: false,
    signInError: '',

    emails: [],
    emailFetchError: '',
    expandedEmailId: null,
    aiStage: 'idle', // idle | generating | ready | error
    aiSummaryText: '',
    aiError: '',
    suggestedTodos: [],
  },

  setState(patch) {
    const update = typeof patch === 'function' ? patch(this.state) : patch;
    Object.assign(this.state, update);
    this.render();
  },

  // ---------- Microsoft sign-in ----------

  async initMsal() {
    const res = await fetch('/api/config');
    this.msalConfig = await res.json();
    this.googleClientId = this.msalConfig.googleClientId || '';

    if (!this.msalConfig.clientId) {
      this.setState({ signInError: 'Server has no MSAL_CLIENT_ID configured yet (.env).' });
      return;
    }

    this.msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId: this.msalConfig.clientId,
        authority: `https://login.microsoftonline.com/${this.msalConfig.tenantId}`,
        redirectUri: this.msalConfig.redirectUri,
      },
      cache: { cacheLocation: 'localStorage' },
    });
    await this.msalInstance.initialize();
    await this.msalInstance.handleRedirectPromise();

    const accounts = this.msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      this.account = accounts[0];
      this.authProvider = 'microsoft';
      this.msalInstance.setActiveAccount(this.account);
      await this.loadProfileAndInbox();
      this.setState({ screen: 'home' });
    }
  },

  async signIn() {
    if (!this.msalInstance) {
      this.setState({ signInError: 'Microsoft sign-in is not configured on the server yet.' });
      return;
    }
    this.setState({ signingIn: true, signInError: '' });
    try {
      const result = await this.msalInstance.loginPopup({ scopes: ['User.Read', 'Mail.Read'] });
      this.account = result.account;
      this.authProvider = 'microsoft';
      this.msalInstance.setActiveAccount(this.account);
      await this.loadProfileAndInbox();
      this.setState({ screen: 'home', signingIn: false });
    } catch (err) {
      console.error('Sign-in failed:', err);
      this.setState({ signingIn: false, signInError: err.message || 'Sign-in failed.' });
    }
  },

  async getGraphToken(scopes) {
    try {
      const result = await this.msalInstance.acquireTokenSilent({ scopes, account: this.account });
      return result.accessToken;
    } catch (err) {
      const result = await this.msalInstance.acquireTokenPopup({ scopes });
      return result.accessToken;
    }
  },

  async loadProfileAndInbox() {
    try {
      const token = await this.getGraphToken(['User.Read']);
      const meRes = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: { Authorization: `Bearer ${token}` },
      });
      const me = await meRes.json();
      const name = me.displayName || this.account.name || this.account.username || 'Signed in';
      this.state.userName = name;
      this.state.userInitial = (name.trim()[0] || '?').toUpperCase();
    } catch (err) {
      console.error('Failed to load profile:', err);
      this.state.userName = this.account.name || this.account.username || 'Signed in';
      this.state.userInitial = (this.state.userName.trim()[0] || '?').toUpperCase();
    }

    // Hydrate from the server's durable copy of this account's data — this is
    // what survives a cleared browser or a switch to a different device,
    // which plain localStorage never could. Local data (e.g. from before
    // this feature existed) stays in place if the server has nothing yet.
    await this.loadUserDataFromServer();

    try {
      const mailToken = await this.getGraphToken(['Mail.Read']);
      const now = new Date();
      const startOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() - now.getDay());
      const filter = `receivedDateTime ge ${startOfWeek.toISOString()}`;
      const mailRes = await fetch(
        'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages' +
          `?$top=50&$orderby=receivedDateTime desc&$filter=${encodeURIComponent(filter)}` +
          '&$select=id,subject,from,bodyPreview,body,receivedDateTime,isRead',
        { headers: { Authorization: `Bearer ${mailToken}`, Prefer: 'outlook.body-content-type="text"' } }
      );
      if (!mailRes.ok) throw new Error(`Graph mail request failed (${mailRes.status})`);
      const mailData = await mailRes.json();
      this.state.emails = (mailData.value || []).map((m) => {
        const senderName = (m.from && m.from.emailAddress && m.from.emailAddress.name) || 'Unknown sender';
        return {
          id: m.id,
          sender: senderName,
          subject: m.subject || '(no subject)',
          snippet: m.bodyPreview || '',
          body: cleanOutlookPlainText((m.body && m.body.content) || m.bodyPreview || ''),
          receivedAt: new Date(m.receivedDateTime),
          time: new Date(m.receivedDateTime).toLocaleString(undefined, { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' }),
          isRead: m.isRead !== false,
          initial: (senderName.trim()[0] || '?').toUpperCase(),
          avatarColor: colorForName(senderName),
        };
      });
      this.state.emailFetchError = '';
    } catch (err) {
      console.error('Failed to load inbox:', err);
      this.state.emailFetchError = 'Could not read your inbox (' + err.message + '). Grant "Mail.Read" consent to your account, or check with your admin.';
      this.state.emails = [];
    }
  },

  // ---------- Google sign-in (Gmail) ----------

  async signInWithGoogle() {
    if (!this.googleClientId) {
      this.setState({ signInError: 'Server has no GOOGLE_CLIENT_ID configured yet (.env).' });
      return;
    }
    if (!window.google || !google.accounts || !google.accounts.oauth2) {
      this.setState({ signInError: 'Google sign-in is still loading — try again in a moment.' });
      return;
    }
    this.setState({ signingIn: true, signInError: '' });
    const tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: this.googleClientId,
      scope: [
        'https://www.googleapis.com/auth/gmail.readonly',
        'https://www.googleapis.com/auth/userinfo.profile',
        'https://www.googleapis.com/auth/userinfo.email',
      ].join(' '),
      callback: async (response) => {
        if (response.error) {
          this.setState({ signingIn: false, signInError: response.error_description || 'Google sign-in failed.' });
          return;
        }
        this.googleAccessToken = response.access_token;
        this.authProvider = 'google';
        await this.loadGoogleProfileAndInbox();
        this.setState({ screen: 'home', signingIn: false });
      },
    });
    tokenClient.requestAccessToken();
  },

  async loadGoogleProfileAndInbox() {
    try {
      const meRes = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
        headers: { Authorization: `Bearer ${this.googleAccessToken}` },
      });
      const me = await meRes.json();
      const name = me.name || me.email || 'Signed in';
      this.state.userName = name;
      this.state.userInitial = (name.trim()[0] || '?').toUpperCase();
    } catch (err) {
      console.error('Failed to load Google profile:', err);
      this.state.userName = 'Signed in';
      this.state.userInitial = '?';
    }

    await this.loadUserDataFromServer();

    try {
      const now = new Date();
      const startOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() - now.getDay());
      const afterSeconds = Math.floor(startOfWeek.getTime() / 1000);
      const listRes = await fetch(
        `https://gmail.googleapis.com/gmail/v1/users/me/messages?maxResults=50&q=${encodeURIComponent(`in:inbox after:${afterSeconds}`)}`,
        { headers: { Authorization: `Bearer ${this.googleAccessToken}` } }
      );
      if (!listRes.ok) throw new Error(`Gmail list request failed (${listRes.status})`);
      const listData = await listRes.json();
      const ids = (listData.messages || []).map((m) => m.id);
      const messages = await Promise.all(ids.map(async (id) => {
        const msgRes = await fetch(
          `https://gmail.googleapis.com/gmail/v1/users/me/messages/${id}?format=full`,
          { headers: { Authorization: `Bearer ${this.googleAccessToken}` } }
        );
        return msgRes.ok ? msgRes.json() : null;
      }));
      this.state.emails = messages.filter(Boolean).map((m) => {
        const headers = (m.payload && m.payload.headers) || [];
        const getHeader = (name) => (headers.find((h) => h.name.toLowerCase() === name.toLowerCase()) || {}).value || '';
        const fromHeader = getHeader('From');
        const senderMatch = fromHeader.match(/^"?([^"<]*)"?\s*<?([^>]*)>?$/);
        const senderName = (senderMatch && senderMatch[1].trim()) || fromHeader || 'Unknown sender';
        const receivedAt = new Date(Number(m.internalDate));
        const isRead = !(m.labelIds || []).includes('UNREAD');
        return {
          id: m.id,
          sender: senderName,
          subject: getHeader('Subject') || '(no subject)',
          snippet: m.snippet || '',
          body: extractGmailBody(m.payload) || m.snippet || '',
          receivedAt,
          time: receivedAt.toLocaleString(undefined, { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' }),
          isRead,
          initial: (senderName.trim()[0] || '?').toUpperCase(),
          avatarColor: colorForName(senderName),
        };
      }).sort((a, b) => b.receivedAt - a.receivedAt);
      this.state.emailFetchError = '';
    } catch (err) {
      console.error('Failed to load Gmail inbox:', err);
      this.state.emailFetchError = 'Could not read your Gmail inbox (' + err.message + ').';
      this.state.emails = [];
    }
  },

  // ---------- Server-side data sync (durable, per-account storage) ----------

  collectUserData() {
    return {
      tasks: this.state.tasks,
      categories: this.state.categories,
      events: this.state.events,
      reminders: this.state.reminders,
      habits: this.habits,
      habitLog: this.habitLog,
      moodLog: this.moodLog,
      taskHistory: this.taskHistory,
      history: this.history,
      emailNotes: this.emailNotes,
    };
  },
  applyUserData(data) {
    if (!data) return;
    if (Array.isArray(data.tasks)) this.state.tasks = data.tasks;
    if (Array.isArray(data.categories) && data.categories.length > 0) this.state.categories = data.categories;
    if (Array.isArray(data.events)) this.state.events = data.events;
    if (Array.isArray(data.reminders)) this.state.reminders = data.reminders;
    if (Array.isArray(data.habits)) this.habits = data.habits;
    if (data.habitLog && typeof data.habitLog === 'object') this.habitLog = data.habitLog;
    if (data.moodLog && typeof data.moodLog === 'object') this.moodLog = data.moodLog;
    if (data.taskHistory && typeof data.taskHistory === 'object') this.taskHistory = data.taskHistory;
    if (data.history && typeof data.history === 'object') this.history = data.history;
    if (data.emailNotes && typeof data.emailNotes === 'object') this.emailNotes = data.emailNotes;
    // Keep localStorage's local cache consistent with whatever the server
    // just handed us, so an offline reload still sees the same data.
    this.saveTasks();
    this.saveCategories();
    this.saveEvents();
    this.saveReminders();
    this.saveHabits();
    this.saveHabitLog();
    this.saveMoodLog();
    this.saveTaskHistory();
    this.saveStats();
    this.saveEmailNotes();
  },
  // Works for whichever provider is currently signed in — the server
  // verifies the token against that same provider, so this can't be used
  // to spoof another account.
  async getSyncAuthToken() {
    if (this.authProvider === 'google') return { token: this.googleAccessToken, provider: 'google' };
    return { token: await this.getGraphToken(['User.Read']), provider: 'microsoft' };
  },
  async loadUserDataFromServer() {
    try {
      const { token, provider } = await this.getSyncAuthToken();
      if (!token) throw new Error('No auth token available');
      const res = await fetch('/api/user-data', { headers: { Authorization: `Bearer ${token}`, 'X-Auth-Provider': provider } });
      if (!res.ok) throw new Error(`Server returned ${res.status}`);
      const { data } = await res.json();
      this.applyUserData(data);
    } catch (err) {
      // No durable copy yet (first time signing in), or the server was
      // unreachable — keep going on whatever's in localStorage already.
      console.error('Could not load saved data from server:', err);
    }
  },
  async saveUserDataToServer() {
    try {
      const { token, provider } = await this.getSyncAuthToken();
      if (!token) throw new Error('No auth token available');
      const res = await fetch('/api/user-data', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}`, 'X-Auth-Provider': provider },
        body: JSON.stringify({ data: this.collectUserData() }),
      });
      if (!res.ok) throw new Error(`Server returned ${res.status}`);
      this._lastServerSaveAt = Date.now();
    } catch (err) {
      // Leave _lastServerSaveAt where it was so the next render's throttle
      // check allows a prompt retry instead of waiting out the full window.
      console.error('Could not save data to server:', err);
    }
  },
  // Time-based rather than debounced: render() fires every second while a
  // focus session is ticking, and a debounce would keep getting reset and
  // never actually fire. This guarantees a save at least every ~10s.
  maybeSyncToServer() {
    if (this._serverSyncInFlight) return;
    if (Date.now() - (this._lastServerSaveAt || 0) < 10000) return;
    this._serverSyncInFlight = true;
    this.saveUserDataToServer().finally(() => { this._serverSyncInFlight = false; });
  },
  syncToServerNow() {
    this.saveUserDataToServer();
  },

  async signOut() {
    if (this.authProvider) await this.saveUserDataToServer();
    try {
      if (this.authProvider === 'google' && this.googleAccessToken && window.google) {
        google.accounts.oauth2.revoke(this.googleAccessToken);
      } else if (this.msalInstance && this.account) {
        await this.msalInstance.logoutPopup({ account: this.account });
      }
    } catch (err) {
      console.error('Sign-out warning:', err);
    }
    this.account = null;
    this.googleAccessToken = null;
    this.authProvider = null;
    this.setState({
      screen: 'onboarding',
      userName: '',
      userInitial: '',
      emails: [],
      emailFetchError: '',
      aiStage: 'idle',
      aiSummaryText: '',
      suggestedTodos: [],
    });
  },

  // ---------- Real AI summary (Claude, via server proxy) ----------

  async generateSummary() {
    const last24hCutoff = new Date(Date.now() - 24 * 60 * 60 * 1000);
    const recentEmails = this.state.emails.filter((e) => e.receivedAt >= last24hCutoff);
    if (recentEmails.length === 0) {
      this.setState({ aiStage: 'error', aiError: 'No emails from the last 24 hours to summarize.' });
      return;
    }
    this.setState({ aiStage: 'generating' });
    try {
      const res = await fetch('/api/summarize', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          emails: recentEmails.map((e) => ({ sender: e.sender, subject: e.subject, snippet: e.snippet })),
        }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Request failed');
      this.setState({
        aiStage: 'ready',
        aiSummaryText: data.summary,
        suggestedTodos: (data.todos || []).map((t, i) => ({ id: 5000 + i, name: t.name, source: t.source, date: t.date || null, added: false })),
      });
    } catch (err) {
      console.error('Summary generation failed:', err);
      this.setState({ aiStage: 'error', aiError: err.message || 'Something went wrong.' });
    }
  },

  addTodoToTasks(id) {
    const todo = this.state.suggestedTodos.find((t) => t.id === id);
    if (!todo || todo.added) return;
    const nextId = Math.max(0, ...this.state.tasks.map((t) => t.id)) + 1;
    todo.added = true;
    this.state.tasks.push({ id: nextId, name: todo.name, done: false, tag: 'Mail', subtasks: [], date: todo.date || null });
    this.render();
  },

  // ---------- Calendar (events + AI email scan) ----------

  selectCalendarDate(dateStr) { this.setState({ calendarSelectedDate: dateStr }); },
  shiftCalendarDate(deltaDays) {
    const d = parseLocalDate(this.state.calendarSelectedDate);
    d.setDate(d.getDate() + deltaDays);
    this.setState({ calendarSelectedDate: dayKey(d) });
  },
  setCalendarViewMode(mode) { this.setState({ calendarViewMode: mode }); },
  goToDayFromCalendar(dateStr) {
    this.setState({ calendarSelectedDate: dateStr, calendarViewMode: 'day' });
  },
  calendarShiftMonth(delta) {
    this.setState((s) => {
      let { year, month } = s.calendarViewMonth;
      month += delta;
      if (month < 0) { month = 11; year -= 1; }
      if (month > 11) { month = 0; year += 1; }
      return { calendarViewMonth: { year, month } };
    });
  },
  calendarGoToMonth(year, month) {
    this.setState({ calendarViewMonth: { year, month }, calendarViewMode: 'month' });
  },
  calendarShiftYear(delta) {
    this.setState((s) => ({ calendarViewYear: s.calendarViewYear + delta }));
  },

  addEvent() {
    const nameInput = document.getElementById('newEventName');
    const name = (nameInput && nameInput.value || '').trim();
    if (!name) return;
    const dateInput = document.getElementById('newEventDate');
    const date = (dateInput && dateInput.value) || this.state.calendarSelectedDate;
    const endDateInput = document.getElementById('newEventEndDate');
    const endDateValue = (endDateInput && endDateInput.value) || null;
    const endDate = endDateValue && endDateValue > date ? endDateValue : null;
    const timeInput = document.getElementById('newEventTime');
    const time = (timeInput && timeInput.value) || null;
    const importantInput = document.getElementById('newEventImportant');
    const important = !!(importantInput && importantInput.checked);
    const nextId = Math.max(0, ...this.state.events.map((e) => e.id)) + 1;
    this.state.events.push({ id: nextId, name, date, endDate, time, important, source: 'manual' });
    nameInput.value = '';
    if (endDateInput) endDateInput.value = '';
    if (timeInput) timeInput.value = '';
    if (importantInput) importantInput.checked = false;
    this.render();
  },
  eventOccursOnDate(e, dateStr) {
    return dateStr >= e.date && dateStr <= (e.endDate || e.date);
  },
  removeEvent(id) {
    this.state.events = this.state.events.filter((e) => e.id !== id);
    this.render();
  },

  async scanEmailsForEvents() {
    if (this.state.emails.length === 0) {
      this.setState({ eventScanStage: 'error', eventScanError: 'No emails loaded yet — open the Inbox first.' });
      return;
    }
    this.setState({ eventScanStage: 'scanning' });
    try {
      const res = await fetch('/api/extract-events', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          emails: this.state.emails.map((e) => ({ sender: e.sender, subject: e.subject, snippet: e.snippet })),
        }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Request failed');
      this.setState({
        eventScanStage: 'ready',
        proposedEvents: (data.events || []).map((e, i) => ({ id: 9000 + i, name: e.name, date: e.date, source: e.source })),
      });
    } catch (err) {
      console.error('Event scan failed:', err);
      this.setState({ eventScanStage: 'error', eventScanError: err.message || 'Something went wrong.' });
    }
  },
  acceptProposedEvent(id) {
    const proposed = this.state.proposedEvents.find((e) => e.id === id);
    if (!proposed) return;
    const nextId = Math.max(0, ...this.state.events.map((e) => e.id)) + 1;
    this.state.events.push({ id: nextId, name: proposed.name, date: proposed.date, endDate: null, time: null, important: false, source: 'ai' });
    this.state.proposedEvents = this.state.proposedEvents.filter((e) => e.id !== id);
    this.render();
  },
  dismissProposedEvent(id) {
    this.state.proposedEvents = this.state.proposedEvents.filter((e) => e.id !== id);
    this.render();
  },

  // ---------- Reminders ----------

  createReminder(date, text) {
    if (!text) return;
    const nextId = Math.max(0, ...this.state.reminders.map((r) => r.id)) + 1;
    this.state.reminders.push({ id: nextId, date, text });
    this.render();
  },
  addReminderForSelectedDate() {
    const input = document.getElementById('newReminderText');
    const text = (input && input.value || '').trim();
    this.createReminder(this.state.calendarSelectedDate, text);
    if (input) input.value = '';
  },
  addTodayReminder() {
    const input = document.getElementById('newTodayReminderText');
    const text = (input && input.value || '').trim();
    this.createReminder(dayKey(new Date()), text);
    if (input) input.value = '';
  },
  handleReminderKeyDown(e, fnName) { if (e.key === 'Enter') this[fnName](); },
  removeReminder(id) {
    this.state.reminders = this.state.reminders.filter((r) => r.id !== id);
    this.render();
  },
  startEditReminder(id) { this.setState({ editingReminderId: id }); },
  saveReminderText(id, inputId) {
    const input = document.getElementById(inputId);
    const text = (input && input.value || '').trim();
    const r = this.state.reminders.find((x) => x.id === id);
    if (r && text) r.text = text;
    this.setState({ editingReminderId: null });
  },
  cancelEditReminder() { this.setState({ editingReminderId: null }); },
  handleEditReminderKeyDown(e, id, inputId) {
    if (e.key === 'Enter') this.saveReminderText(id, inputId);
    else if (e.key === 'Escape') this.cancelEditReminder();
  },

  // ---------- Navigation ----------

  backToHome() { this.setState({ screen: 'home' }); },
  goToSetup() { this.setState({ screen: 'setup' }); },
  goToStats() { this.setState({ screen: 'stats' }); },
  goToPrevMonth() {
    this.setState((s) => {
      let { year, month } = s.statsViewMonth;
      month -= 1;
      if (month < 0) { month = 11; year -= 1; }
      return { statsViewMonth: { year, month } };
    });
  },
  goToNextMonth() {
    this.setState((s) => {
      let { year, month } = s.statsViewMonth;
      month += 1;
      if (month > 11) { month = 0; year += 1; }
      return { statsViewMonth: { year, month } };
    });
  },
  selectStatsDate(key) { this.setState({ statsSelectedDate: key }); },
  goToInbox() { this.setState({ screen: 'inbox' }); },
  goToCalendar() { this.setState({ screen: 'calendar', calendarSelectedDate: dayKey(new Date()) }); },
  goToSettings() { this.setState({ screen: 'settings' }); },

  // ---------- Settings ----------

  selectTheme(theme) { this.setState({ theme }); },
  toggleSetting(key) {
    this.state.settingsToggles[key] = !this.state.settingsToggles[key];
    this.render();
  },

  // ---------- Task categories ----------

  loadCategories() {
    try {
      const saved = JSON.parse(localStorage.getItem('focusDayCategories') || 'null');
      if (Array.isArray(saved) && saved.length > 0) this.state.categories = saved;
    } catch {
      // keep the default categories
    }
  },
  saveCategories() {
    localStorage.setItem('focusDayCategories', JSON.stringify(this.state.categories));
  },

  // ---------- Tasks, events, reminders (persisted — the actual daily data) ----------

  loadTasks() {
    try {
      const saved = JSON.parse(localStorage.getItem('focusDayTasks') || 'null');
      if (Array.isArray(saved)) this.state.tasks = saved;
    } catch {
      // keep the default (empty) task list
    }
  },
  saveTasks() {
    localStorage.setItem('focusDayTasks', JSON.stringify(this.state.tasks));
  },
  loadEvents() {
    try {
      const saved = JSON.parse(localStorage.getItem('focusDayEvents') || 'null');
      if (Array.isArray(saved)) this.state.events = saved;
    } catch {
      // keep the default (empty) events list
    }
  },
  saveEvents() {
    localStorage.setItem('focusDayEvents', JSON.stringify(this.state.events));
  },
  loadReminders() {
    try {
      const saved = JSON.parse(localStorage.getItem('focusDayReminders') || 'null');
      if (Array.isArray(saved)) this.state.reminders = saved;
    } catch {
      // keep the default (empty) reminders list
    }
  },
  saveReminders() {
    localStorage.setItem('focusDayReminders', JSON.stringify(this.state.reminders));
  },
  addCategory() {
    const input = document.getElementById('newCategoryInput');
    const name = (input && input.value || '').trim();
    if (!name || this.state.categories.includes(name)) return;
    this.state.categories.push(name);
    this.saveCategories();
    this.render();
  },
  handleCategoryKeyDown(e) { if (e.key === 'Enter') this.addCategory(); },
  renameCategory(index, newName) {
    const trimmed = (newName || '').trim();
    const oldName = this.state.categories[index];
    if (!trimmed || trimmed === oldName) { this.render(); return; }
    this.state.categories[index] = trimmed;
    // Existing tasks tagged with the old name follow the rename.
    this.state.tasks.forEach((t) => { if (t.tag === oldName) t.tag = trimmed; });
    this.saveCategories();
    this.render();
  },
  removeCategory(index) {
    if (this.state.categories.length <= 1) return;
    this.state.categories.splice(index, 1);
    this.saveCategories();
    this.render();
  },

  // ---------- Tasks ----------

  selectTask(id) { this.setState({ selectedTaskId: id }); },
  toggleTaskDone(id) {
    const t = this.state.tasks.find((x) => x.id === id);
    if (t) {
      t.done = !t.done;
      if (t.done) this.logTaskCompleted(t);
    }
    this.render();
  },
  removeTask(id) {
    this.state.tasks = this.state.tasks.filter((t) => t.id !== id);
    this.render();
  },
  addTask() {
    const input = document.getElementById('newTaskInput');
    const name = (input && input.value || '').trim();
    if (!name) return;
    const categorySelect = document.getElementById('newTaskCategory');
    const tag = (categorySelect && categorySelect.value) || this.state.categories[0];
    const dateInput = document.getElementById('newTaskDate');
    const date = (dateInput && dateInput.value) || null;
    const nextId = Math.max(0, ...this.state.tasks.map((t) => t.id)) + 1;
    const subtasks = this.state.newTaskDraftSubtasks.map((n, i) => ({ id: i + 1, name: n, done: false, date: null }));
    this.state.tasks.push({ id: nextId, name, done: false, tag, subtasks, date });
    this.state.newTaskDraftSubtasks = [];
    this.render();
  },
  setTaskDate(taskId, value) {
    const t = this.state.tasks.find((x) => x.id === taskId);
    if (!t) return;
    t.date = value || null;
    this.setState({ datePickerTaskId: null });
  },
  setSubtaskDate(taskId, subId, value) {
    const t = this.state.tasks.find((x) => x.id === taskId);
    const st = t && t.subtasks && t.subtasks.find((s) => s.id === subId);
    if (!st) return;
    st.date = value || null;
    this.setState({ datePickerSubtask: null });
  },
  toggleDatePicker(taskId) {
    this.setState((s) => ({ datePickerTaskId: s.datePickerTaskId === taskId ? null : taskId }));
  },
  toggleSubtaskDatePicker(taskId, subId) {
    this.setState((s) => ({
      datePickerSubtask: s.datePickerSubtask && s.datePickerSubtask.taskId === taskId && s.datePickerSubtask.subId === subId
        ? null : { taskId, subId },
    }));
  },

  startEditTask(taskId) { this.setState({ editingTaskId: taskId }); },
  saveTaskName(taskId, inputId) {
    const input = document.getElementById(inputId);
    const name = (input && input.value || '').trim();
    const t = this.state.tasks.find((x) => x.id === taskId);
    if (t && name) t.name = name;
    this.setState({ editingTaskId: null });
  },
  cancelEditTask() { this.setState({ editingTaskId: null }); },
  handleEditTaskKeyDown(e, taskId, inputId) {
    if (e.key === 'Enter') this.saveTaskName(taskId, inputId);
    else if (e.key === 'Escape') this.cancelEditTask();
  },

  startEditSubtask(taskId, subId) { this.setState({ editingSubtask: { taskId, subId } }); },
  saveSubtaskName(taskId, subId, inputId) {
    const input = document.getElementById(inputId);
    const name = (input && input.value || '').trim();
    const t = this.state.tasks.find((x) => x.id === taskId);
    const st = t && t.subtasks && t.subtasks.find((s) => s.id === subId);
    if (st && name) st.name = name;
    this.setState({ editingSubtask: null });
  },
  cancelEditSubtask() { this.setState({ editingSubtask: null }); },
  handleEditSubtaskKeyDown(e, taskId, subId, inputId) {
    if (e.key === 'Enter') this.saveSubtaskName(taskId, subId, inputId);
    else if (e.key === 'Escape') this.cancelEditSubtask();
  },
  handleTaskKeyDown(e) { if (e.key === 'Enter') this.addTask(); },
  addDraftSubtask() {
    const input = document.getElementById('newTaskDraftSubtaskInput');
    const name = (input && input.value || '').trim();
    if (!name) return;
    this.state.newTaskDraftSubtasks.push(name);
    this.render();
  },
  handleDraftSubtaskKeyDown(e) { if (e.key === 'Enter') { e.preventDefault(); this.addDraftSubtask(); } },
  removeDraftSubtask(index) {
    this.state.newTaskDraftSubtasks.splice(index, 1);
    this.render();
  },

  removeSetupTask(id) {
    this.state.tasks = this.state.tasks.filter((t) => t.id !== id);
    if (this.state.selectedTaskId === id) {
      this.state.selectedTaskId = this.state.tasks[0] ? this.state.tasks[0].id : null;
    }
    this.render();
  },
  addSetupTask() {
    const input = document.getElementById('newSetupTaskInput');
    const name = (input && input.value || '').trim();
    if (!name) return;
    const categorySelect = document.getElementById('newSetupTaskCategory');
    const tag = (categorySelect && categorySelect.value) || this.state.categories[0];
    const nextId = Math.max(0, ...this.state.tasks.map((t) => t.id)) + 1;
    const subtasks = this.state.newSetupTaskDraftSubtasks.map((n, i) => ({ id: i + 1, name: n, done: false, date: null }));
    this.state.tasks.push({ id: nextId, name, done: false, tag, subtasks, date: null });
    this.state.selectedTaskId = nextId;
    this.state.newSetupTaskDraftSubtasks = [];
    this.render();
  },
  handleSetupTaskKeyDown(e) { if (e.key === 'Enter') this.addSetupTask(); },
  addSetupDraftSubtask() {
    const input = document.getElementById('newSetupTaskDraftSubtaskInput');
    const name = (input && input.value || '').trim();
    if (!name) return;
    this.state.newSetupTaskDraftSubtasks.push(name);
    this.render();
  },
  handleSetupDraftSubtaskKeyDown(e) { if (e.key === 'Enter') { e.preventDefault(); this.addSetupDraftSubtask(); } },
  removeSetupDraftSubtask(index) {
    this.state.newSetupTaskDraftSubtasks.splice(index, 1);
    this.render();
  },
  focusOnTask(id) { this.setState({ selectedTaskId: id, screen: 'setup' }); },

  // ---------- Subtasks (breaking a task into smaller steps) ----------

  toggleSubtaskInput(taskId) {
    this.setState((s) => ({ expandedSubtaskTaskId: s.expandedSubtaskTaskId === taskId ? null : taskId }));
  },
  addSubtaskToTask(taskId, inputId) {
    const input = document.getElementById(inputId);
    const name = (input && input.value || '').trim();
    if (!name) return;
    const t = this.state.tasks.find((x) => x.id === taskId);
    if (!t) return;
    if (!t.subtasks) t.subtasks = [];
    const nextId = Math.max(0, ...t.subtasks.map((st) => st.id)) + 1;
    t.subtasks.push({ id: nextId, name, done: false, date: null });
    this.render();
  },
  handleSubtaskKeyDown(e, taskId, inputId) { if (e.key === 'Enter') this.addSubtaskToTask(taskId, inputId); },
  removeSubtask(taskId, subId) {
    const t = this.state.tasks.find((x) => x.id === taskId);
    if (t && t.subtasks) t.subtasks = t.subtasks.filter((st) => st.id !== subId);
    this.render();
  },
  toggleSubtaskDone(taskId, subId) {
    const t = this.state.tasks.find((x) => x.id === taskId);
    const st = t && t.subtasks && t.subtasks.find((s) => s.id === subId);
    if (st) st.done = !st.done;
    this.render();
  },

  // ---------- Real stats persistence (localStorage) ----------

  loadStats() {
    try {
      this.history = JSON.parse(localStorage.getItem('focusDayHistory') || '{}');
    } catch {
      this.history = {};
    }
  },
  saveStats() {
    localStorage.setItem('focusDayHistory', JSON.stringify(this.history));
  },
  todayRecord() {
    return this.history[dayKey(new Date())] || { focusSeconds: 0, sessionsCompleted: 0 };
  },
  // Called whenever a focus phase actually ends (fully or early) — this is the
  // only place real focus time gets banked, so stats always reflect time truly spent.
  recordFocusCompletion(seconds, completedFully) {
    const key = dayKey(new Date());
    const rec = this.history[key] || { focusSeconds: 0, sessionsCompleted: 0 };
    rec.focusSeconds += seconds;
    if (completedFully) rec.sessionsCompleted += 1;
    this.history[key] = rec;
    this.saveStats();
  },
  historyForLastNDays(n) {
    const days = [];
    for (let i = n - 1; i >= 0; i--) {
      const d = new Date();
      d.setDate(d.getDate() - i);
      const rec = this.history[dayKey(d)] || { focusSeconds: 0, sessionsCompleted: 0 };
      days.push({ date: d, ...rec });
    }
    return days;
  },
  allTimeTotals() {
    let focusSeconds = 0;
    let sessions = 0;
    Object.values(this.history).forEach((r) => { focusSeconds += r.focusSeconds; sessions += r.sessionsCompleted; });
    return { focusSeconds, sessions };
  },
  currentStreak() {
    let streak = 0;
    for (let i = 0; ; i++) {
      const d = new Date();
      d.setDate(d.getDate() - i);
      const rec = this.history[dayKey(d)];
      if (rec && rec.focusSeconds > 0) streak++;
      else break;
    }
    return streak;
  },
  longestStreak() {
    const keys = Object.keys(this.history).filter((k) => this.history[k].focusSeconds > 0).sort();
    if (keys.length === 0) return 0;
    let longest = 1;
    let current = 1;
    for (let i = 1; i < keys.length; i++) {
      const diffDays = Math.round((new Date(keys[i]) - new Date(keys[i - 1])) / 86400000);
      current = diffDays === 1 ? current + 1 : 1;
      longest = Math.max(longest, current);
    }
    return longest;
  },

  // ---------- Completed-task log (for History + Stats day detail) ----------

  loadTaskHistory() {
    try {
      this.taskHistory = JSON.parse(localStorage.getItem('focusDayTaskHistory') || '{}');
    } catch {
      this.taskHistory = {};
    }
  },
  saveTaskHistory() {
    localStorage.setItem('focusDayTaskHistory', JSON.stringify(this.taskHistory));
  },
  logTaskCompleted(task) {
    const key = dayKey(new Date());
    if (!this.taskHistory[key]) this.taskHistory[key] = [];
    this.taskHistory[key].push({ name: task.name, tag: task.tag });
    this.saveTaskHistory();
  },

  // ---------- Mood log (anytime check-ins, for Stats day detail) ----------

  loadMoodLog() {
    try {
      this.moodLog = JSON.parse(localStorage.getItem('focusDayMoodLog') || '{}');
    } catch {
      this.moodLog = {};
    }
  },
  saveMoodLog() {
    localStorage.setItem('focusDayMoodLog', JSON.stringify(this.moodLog));
  },
  logMood(mood) {
    const key = dayKey(new Date());
    if (!this.moodLog[key]) this.moodLog[key] = [];
    this.moodLog[key].push({ time: Date.now(), mood });
    this.saveMoodLog();
    this.setState({ lastLoggedMood: mood });
  },

  // ---------- Per-email notes ----------

  loadEmailNotes() {
    try {
      this.emailNotes = JSON.parse(localStorage.getItem('focusDayEmailNotes') || '{}');
    } catch {
      this.emailNotes = {};
    }
  },
  saveEmailNotes() {
    localStorage.setItem('focusDayEmailNotes', JSON.stringify(this.emailNotes));
  },
  setEmailNote(emailId, text) {
    if (text) this.emailNotes[emailId] = text;
    else delete this.emailNotes[emailId];
    this.saveEmailNotes();
  },
  toggleEmailExpanded(emailId) {
    this.setState((s) => ({ expandedEmailId: s.expandedEmailId === emailId ? null : emailId }));
  },

  // ---------- Habits ----------

  loadHabits() {
    try {
      this.habits = JSON.parse(localStorage.getItem('focusDayHabits') || '[]');
    } catch {
      this.habits = [];
    }
  },
  saveHabits() {
    localStorage.setItem('focusDayHabits', JSON.stringify(this.habits));
  },
  loadHabitLog() {
    try {
      this.habitLog = JSON.parse(localStorage.getItem('focusDayHabitLog') || '{}');
    } catch {
      this.habitLog = {};
    }
  },
  saveHabitLog() {
    localStorage.setItem('focusDayHabitLog', JSON.stringify(this.habitLog));
  },
  selectNewHabitIcon(icon) { this.setState({ newHabitIcon: icon }); },
  setNewHabitTarget(target) { this.setState({ newHabitTarget: target }); },
  toggleNewHabitForm() { this.setState((s) => ({ showNewHabitForm: !s.showNewHabitForm })); },
  addHabit() {
    const input = document.getElementById('newHabitName');
    const name = (input && input.value || '').trim();
    if (!name) return;
    const nextId = Math.max(0, ...this.habits.map((h) => h.id)) + 1;
    this.habits.push({ id: nextId, name, icon: this.state.newHabitIcon, targetPerWeek: this.state.newHabitTarget });
    this.saveHabits();
    input.value = '';
    this.setState({ newHabitIcon: HABIT_ICONS[0], newHabitTarget: 7, showNewHabitForm: false });
  },
  removeHabit(id) {
    this.habits = this.habits.filter((h) => h.id !== id);
    this.saveHabits();
    this.render();
  },
  toggleHabitToday(habitId) {
    const key = dayKey(new Date());
    if (!this.habitLog[key]) this.habitLog[key] = {};
    if (this.habitLog[key][habitId]) delete this.habitLog[key][habitId];
    else this.habitLog[key][habitId] = true;
    this.saveHabitLog();
    this.render();
  },
  habitStreak(habitId) {
    let streak = 0;
    for (let i = 0; ; i++) {
      const d = new Date();
      d.setDate(d.getDate() - i);
      const key = dayKey(d);
      if (this.habitLog[key] && this.habitLog[key][habitId]) streak++;
      else break;
    }
    return streak;
  },
  checkInsThisWeek(habitId) {
    const now = new Date();
    const startOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() - now.getDay());
    let count = 0;
    for (let i = 0; i <= now.getDay(); i++) {
      const d = new Date(startOfWeek);
      d.setDate(d.getDate() + i);
      const key = dayKey(d);
      if (this.habitLog[key] && this.habitLog[key][habitId]) count++;
    }
    return count;
  },

  // ---------- Focus timer (focus / break / cycles) ----------

  toggleBreakEnabled() { this.setState((s) => ({ breakEnabled: !s.breakEnabled })); },
  changeCycles(delta) {
    const next = Math.max(1, Math.min(10, this.state.cycles + delta));
    this.setState({ cycles: next });
  },

  startSession() {
    if (this.state.interval) clearInterval(this.state.interval);
    this.state.timerTaskId = this.state.selectedTaskId;
    this.state.currentCycleIndex = 0;
    this.state.lastSessionFocusSeconds = 0;
    this.state.lastSessionCycles = 0;
    if (this.state.settingsToggles.autoMode) {
      // Auto mode: no preset length, no breaks/cycles — just count up until
      // the user ends it, and log whatever time actually elapsed.
      this.beginPhase('focus-auto', null);
    } else {
      this.beginPhase('focus', this.state.focusMinutes);
    }
    this.setState({ screen: 'timer' });
  },

  beginPhase(phase, minutes) {
    this.state.currentPhase = phase;
    this.state.phaseDurationMin = minutes;
    this.state.remainingSeconds = minutes != null ? minutes * 60 : 0;
    this.state.phaseElapsedSeconds = 0;
    this.state.isPaused = false;
    this.state.interval = setInterval(() => this.tick(), 1000);
  },

  tick() {
    if (this.state.isPaused) return;
    this.state.phaseElapsedSeconds += 1;
    if (this.state.currentPhase === 'focus-auto') {
      // Counts up with no target — only "End session" stops it.
      this.render();
      return;
    }
    this.state.remainingSeconds -= 1;
    if (this.state.remainingSeconds <= 0) {
      clearInterval(this.state.interval);
      this.advancePhase(true);
      return;
    }
    this.render();
  },

  // Called when a phase ends, either naturally (completedFully=true) or via
  // "End session" (completedFully=false). Only a fully-completed focus phase
  // counts as a finished session — ending early still banks the actual
  // seconds spent, but doesn't falsely credit the whole planned duration.
  // (Auto mode has no target to fall short of, so ending it always counts —
  // see endSession.)
  advancePhase(completedFully) {
    clearInterval(this.state.interval);
    const wasFocus = this.state.currentPhase === 'focus' || this.state.currentPhase === 'focus-auto';
    if (wasFocus) {
      this.state.lastSessionFocusSeconds += this.state.phaseElapsedSeconds;
      if (completedFully) this.state.lastSessionCycles += 1;
      // Bank real elapsed seconds the moment the phase ends, whether it ran to
      // completion or was cut short — this is what stats are read from.
      this.recordFocusCompletion(this.state.phaseElapsedSeconds, completedFully);
      // Every time a focus phase ends, check in on the task before deciding
      // what happens next (break / next cycle / done).
      this.state.checkinCompletedFully = completedFully;
      this.setState({ screen: 'checkin' });
      return;
    }

    if (!completedFully) {
      this.finishAllSessions();
      return;
    }
    const isLastCycle = this.state.currentCycleIndex >= this.state.cycles - 1;
    if (!isLastCycle) {
      this.state.currentCycleIndex += 1;
      this.beginPhase('focus', this.state.focusMinutes);
      this.render();
      return;
    }
    this.finishAllSessions();
  },

  // Resumes the focus/break/cycle flow once the check-in (and optional
  // breakdown) screens are done. Ending the session early skips straight to
  // "complete" instead of starting a break or another cycle — and so does
  // auto mode, which never has breaks or cycles to move into.
  proceedAfterFocusPhase() {
    if (!this.state.checkinCompletedFully || this.state.currentPhase === 'focus-auto') {
      this.finishAllSessions();
      return;
    }
    // A cycle is focus + break, so a completed focus phase always moves into
    // its break (if breaks are on) — even on the last cycle.
    if (this.state.breakEnabled) {
      this.beginPhase('break', this.state.breakMinutes);
      this.setState({ screen: 'timer' });
      return;
    }
    const isLastCycle = this.state.currentCycleIndex >= this.state.cycles - 1;
    if (!isLastCycle) {
      this.state.currentCycleIndex += 1;
      this.beginPhase('focus', this.state.focusMinutes);
      this.setState({ screen: 'timer' });
      return;
    }
    this.finishAllSessions();
  },

  // ---------- Post-focus check-in ----------

  checkinTaskDone() {
    const t = this.state.tasks.find((x) => x.id === this.state.timerTaskId);
    if (t) {
      t.done = true;
      this.logTaskCompleted(t);
    }
    this.proceedAfterFocusPhase();
  },
  checkinTaskNotDone() {
    this.setState({ screen: 'breakdown-prompt' });
  },
  declineBreakdown() {
    this.proceedAfterFocusPhase();
  },
  acceptBreakdown() {
    this.setState({ screen: 'breakdown-edit' });
  },
  finishBreakdown() {
    this.proceedAfterFocusPhase();
  },

  skipBreak() {
    if (this.state.currentPhase !== 'break') return;
    clearInterval(this.state.interval);
    const isLastCycle = this.state.currentCycleIndex >= this.state.cycles - 1;
    if (!isLastCycle) {
      this.state.currentCycleIndex += 1;
      this.beginPhase('focus', this.state.focusMinutes);
      this.render();
    } else {
      this.finishAllSessions();
    }
  },

  finishAllSessions() {
    this.state.interval = null;
    this.setState({ screen: 'complete' });
  },

  togglePause() { this.setState((s) => ({ isPaused: !s.isPaused })); },
  endSession() {
    if (this.state.interval) clearInterval(this.state.interval);
    // Auto mode has no target duration to fall short of — ending it is
    // always a genuine completion of whatever time was actually spent.
    const wasAuto = this.state.currentPhase === 'focus-auto';
    this.advancePhase(wasAuto);
  },

  // ---------- Render ----------

  theme() {
    const dark = this.state.theme === 'dark';
    return {
      dark,
      pageBg: dark ? 'oklch(0.18 0.018 75)' : 'oklch(0.97 0.016 85)',
      sidebarBg: dark ? 'oklch(0.15 0.016 75)' : 'oklch(0.975 0.013 85)',
      sidebarBorder: dark ? 'oklch(0.28 0.02 75)' : 'oklch(0.9 0.016 75)',
      cardBg: dark ? 'oklch(0.23 0.02 75)' : 'oklch(0.995 0.007 85)',
      cardBorder: dark ? 'oklch(0.32 0.022 75)' : 'oklch(0.91 0.014 75)',
      divider: dark ? 'oklch(0.3 0.02 75)' : 'oklch(0.92 0.012 75)',
      rowDivider: dark ? 'oklch(0.27 0.018 75)' : 'oklch(0.95 0.009 75)',
      text: dark ? 'oklch(0.93 0.015 85)' : 'oklch(0.24 0.02 90)',
      textMuted: dark ? 'oklch(0.68 0.02 80)' : 'oklch(0.52 0.016 80)',
      hoverBg: dark ? 'oklch(0.27 0.02 75)' : 'oklch(0.95 0.016 80)',
      subtleBg: dark ? 'oklch(0.25 0.02 75)' : 'oklch(0.955 0.016 80)',
      inputBorder: dark ? 'oklch(0.38 0.022 75)' : 'oklch(0.85 0.016 75)',
      unselectedBorder: dark ? 'oklch(0.35 0.022 75)' : 'oklch(0.88 0.014 75)',
      tagBg: dark ? 'oklch(0.28 0.02 75)' : 'oklch(0.94 0.013 75)',
      accentSoftBg: dark ? 'oklch(0.28 0.06 152)' : 'oklch(0.94 0.05 152)',
      accentSoftBorder: dark ? 'oklch(0.4 0.08 152)' : 'oklch(0.82 0.07 152)',
      spinnerTrack: dark ? 'oklch(0.38 0.02 75)' : 'oklch(0.9 0.014 75)',
      navySoftBg: dark ? 'oklch(0.26 0.04 260)' : 'oklch(0.95 0.025 260)',
      navySoftBorder: dark ? 'oklch(0.38 0.06 260)' : 'oklch(0.82 0.05 260)',
      navyText: dark ? 'oklch(0.85 0.05 260)' : 'oklch(0.32 0.08 260)',
      navyIcon: 'oklch(0.32 0.08 260)',
      accentText: dark ? 'oklch(0.75 0.13 152)' : 'oklch(0.36 0.09 152)',
      accentSolid: dark ? 'oklch(0.52 0.15 152)' : 'oklch(0.36 0.09 152)',
      orangeText: dark ? 'oklch(0.75 0.13 50)' : 'oklch(0.6 0.13 50)',
      orangeSoftBg: dark ? 'oklch(0.3 0.06 50)' : 'oklch(0.94 0.05 50)',
      errorSoftBg: dark ? 'oklch(0.28 0.05 30)' : 'oklch(0.96 0.03 30)',
      errorSoftBorder: dark ? 'oklch(0.4 0.08 30)' : 'oklch(0.85 0.06 30)',
      errorText: dark ? 'oklch(0.78 0.13 30)' : 'oklch(0.4 0.13 30)',
      onboardingBg: dark
        ? 'radial-gradient(circle at 30% 20%, oklch(0.24 0.03 152), oklch(0.15 0.018 75) 60%)'
        : 'radial-gradient(circle at 30% 20%, oklch(0.94 0.04 90), oklch(0.97 0.016 85) 60%)',
    };
  },

  renderLogo(sizePx, color) {
    const ringSize = sizePx * 0.62;
    const ringBorder = sizePx * 0.15;
    const dotSize = sizePx * 0.26;
    return `<span style="display:inline-flex;align-items:baseline;font-family:'Manrope',sans-serif;font-weight:800;font-size:${sizePx}px;letter-spacing:-0.01em;color:${color}">af<span style="display:inline-block;width:${ringSize}px;height:${ringSize}px;margin:0 1px;border:${ringBorder}px solid currentColor;border-radius:50%;position:relative;transform:translateY(${sizePx * 0.09}px);flex:none"><span style="position:absolute;top:50%;left:50%;width:${dotSize}px;height:${dotSize}px;background:currentColor;border-radius:50%;transform:translate(-50%,-50%)"></span></span>cusday</span>`;
  },

  renderSidebar(T) {
    const s = this.state;
    const active = (cond) => (cond ? `color:${this.theme().accentText};background:${s.theme === 'dark' ? 'oklch(0.3 0.06 152)' : 'oklch(0.94 0.05 152)'}` : `color:${T.text};background:transparent`);
    const svgWrap = (inner) => `<svg width="18" height="18" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" style="flex:none">${inner}</svg>`;
    const icons = {
      Home: svgWrap('<path d="M3 9.5L10 3.5l7 6"/><path d="M4.5 8.5V16a1 1 0 0 0 1 1h9a1 1 0 0 0 1-1V8.5"/><path d="M8 17v-5h4v5"/>'),
      Focus: svgWrap('<circle cx="10" cy="10" r="6.5"/><circle cx="10" cy="10" r="2.2" fill="currentColor" stroke="none"/>'),
      Inbox: svgWrap('<rect x="2.5" y="4.5" width="15" height="11" rx="1.5"/><path d="M2.5 5.5l7.5 6 7.5-6"/>'),
      Calendar: svgWrap('<rect x="2.5" y="4.5" width="15" height="12.5" rx="1.5"/><path d="M2.5 8.5h15"/><path d="M6.5 2.5v3"/><path d="M13.5 2.5v3"/>'),
      Stats: svgWrap('<path d="M4 16V10"/><path d="M10 16V6"/><path d="M16 16V12"/>'),
      Settings: svgWrap('<line x1="3" y1="6" x2="17" y2="6"/><circle cx="7" cy="6" r="1.8" fill="currentColor" stroke="none"/><line x1="3" y1="10" x2="17" y2="10"/><circle cx="13" cy="10" r="1.8" fill="currentColor" stroke="none"/><line x1="3" y1="14" x2="17" y2="14"/><circle cx="9" cy="14" r="1.8" fill="currentColor" stroke="none"/>'),
    };
    const items = [
      { label: 'Home', cond: s.screen === 'home', fn: 'App.backToHome()' },
      { label: 'Focus', cond: s.screen === 'setup' || s.screen === 'timer', fn: 'App.goToSetup()' },
      { label: 'Inbox', cond: s.screen === 'inbox', fn: 'App.goToInbox()' },
      { label: 'Calendar', cond: s.screen === 'calendar', fn: 'App.goToCalendar()' },
      { label: 'Stats', cond: s.screen === 'stats', fn: 'App.goToStats()' },
      { label: 'Settings', cond: s.screen === 'settings', fn: 'App.goToSettings()' },
    ];
    return `
      <div style="width:220px;flex:none;background:${T.sidebarBg};border-right:1px solid ${T.sidebarBorder};display:flex;flex-direction:column;padding:28px 20px;box-sizing:border-box;${this._animate ? 'animation:fadeIn .3s ease' : ''}">
        <div style="display:flex;align-items:center;margin-bottom:40px;padding-left:4px">
          ${this.renderLogo(24, T.accentText)}
        </div>
        <div style="display:flex;flex-direction:column;gap:2px">
          ${items.map((item) => `
            <div class="nav-item" onclick="${item.fn}" style="display:flex;align-items:center;gap:12px;padding:10px 12px;border-radius:9px;cursor:pointer;font-size:14px;font-weight:500;${active(item.cond)}">
              ${icons[item.label]}
              ${item.label}
            </div>`).join('')}
        </div>
        <div style="margin-top:auto;display:flex;align-items:center;gap:10px;padding:10px 8px;border-radius:10px;background:${T.subtleBg}">
          <div style="width:32px;height:32px;border-radius:50%;background:oklch(0.7 0.12 152);display:flex;align-items:center;justify-content:center;color:white;font-weight:700;font-size:13px;flex:none">${esc(this.state.userInitial)}</div>
          <div style="min-width:0">
            <div style="font-size:12.5px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${esc(this.state.userName)}</div>
            <div style="font-size:11px;color:${T.textMuted}">Microsoft account</div>
          </div>
        </div>
      </div>`;
  },

  renderOnboarding(T) {
    const s = this.state;
    return `
      <div style="flex:1;display:flex;align-items:center;justify-content:center;background:${T.onboardingBg}">
        <div style="width:400px;display:flex;flex-direction:column;align-items:center;text-align:center;animation:popIn .4s cubic-bezier(.2,.8,.2,1)">
          <div style="margin-bottom:20px">${this.renderLogo(30, T.accentText)}</div>
          <div style="font-size:15px;line-height:1.5;color:${T.textMuted};margin-bottom:36px">Plan your day, protect your attention, and track every focused minute — synced with your Outlook or Gmail account.</div>
          <div class="signin-btn" onclick="${s.signingIn ? '' : 'App.signIn()'}" style="width:100%;box-sizing:border-box;display:flex;align-items:center;justify-content:center;gap:10px;padding:13px 20px;background:${T.cardBg};color:${T.text};border:1px solid ${T.cardBorder};border-radius:11px;font-weight:600;font-size:14.5px;cursor:pointer;box-shadow:0 1px 2px rgba(0,0,0,0.04);margin-bottom:10px">
            <svg width="18" height="18" viewBox="0 0 21 21"><rect x="1" y="1" width="9" height="9" fill="#f25022"></rect><rect x="11" y="1" width="9" height="9" fill="#7fba00"></rect><rect x="1" y="11" width="9" height="9" fill="#00a4ef"></rect><rect x="11" y="11" width="9" height="9" fill="#ffb900"></rect></svg>
            ${s.signingIn ? 'Signing in…' : 'Sign in with Microsoft'}
          </div>
          <div class="signin-btn" onclick="${s.signingIn ? '' : 'App.signInWithGoogle()'}" style="width:100%;box-sizing:border-box;display:flex;align-items:center;justify-content:center;gap:10px;padding:13px 20px;background:${T.cardBg};color:${T.text};border:1px solid ${T.cardBorder};border-radius:11px;font-weight:600;font-size:14.5px;cursor:pointer;box-shadow:0 1px 2px rgba(0,0,0,0.04)">
            <svg width="18" height="18" viewBox="0 0 48 48"><path fill="#FFC107" d="M43.6 20.5H42V20H24v8h11.3C33.7 32.7 29.3 36 24 36c-6.6 0-12-5.4-12-12s5.4-12 12-12c3.1 0 5.9 1.2 8 3.1l5.7-5.7C34.5 6.5 29.5 4.5 24 4.5 13.2 4.5 4.5 13.2 4.5 24S13.2 43.5 24 43.5 43.5 34.8 43.5 24c0-1.2-.1-2.4-.4-3.5z"></path><path fill="#FF3D00" d="M6.3 14.7l6.6 4.8C14.7 15.6 19 12.5 24 12.5c3.1 0 5.9 1.2 8 3.1l5.7-5.7C34.5 6.5 29.5 4.5 24 4.5c-7.7 0-14.4 4.4-17.7 10.2z"></path><path fill="#4CAF50" d="M24 43.5c5.4 0 10.3-1.9 14.1-5.4l-6.5-5.5C29.6 34.1 27 35 24 35c-5.3 0-9.7-3.3-11.3-8l-6.6 5.1C9.6 39 16.2 43.5 24 43.5z"></path><path fill="#1976D2" d="M43.6 20.5H42V20H24v8h11.3c-.8 2.2-2.2 4.1-4.1 5.5l6.5 5.5C41.2 36 43.5 30.7 43.5 24c0-1.2-.1-2.4-.4-3.5z"></path></svg>
            ${s.signingIn ? 'Signing in…' : 'Sign in with Google'}
          </div>
          ${s.signInError ? `<div style="font-size:12.5px;color:oklch(0.55 0.15 30);margin-top:14px">${esc(s.signInError)}</div>` : ''}
          <div style="font-size:12px;color:${T.textMuted};margin-top:20px">By continuing you agree to the Terms and Privacy Policy</div>
        </div>
      </div>`;
  },

  renderHome(T) {
    const s = this.state;
    const doneCount = s.tasks.filter((t) => t.done).length;
    const now = new Date();
    const todayKeyStr = dayKey(now);
    const todayLabel = now.toLocaleDateString(undefined, { weekday: 'long', month: 'long', day: 'numeric' });
    const hour = now.getHours();
    const timeGreeting = hour < 12 ? 'Good morning' : hour < 18 ? 'Good afternoon' : 'Good evening';
    const firstName = (s.userName || 'there').split(' ')[0];
    const openTasks = s.tasks.filter((t) => !t.done);
    const todayMoods = (this.moodLog[todayKeyStr] || []).slice().sort((a, b) => a.time - b.time);
    const remindersToday = s.reminders.filter((r) => r.date === todayKeyStr);

    const moodValue = { awful: 1, bad: 2, okay: 3, good: 4, great: 5 };
    const chartW = 300, chartH = 80, padX = 8, padY = 10;
    const moodPoints = todayMoods.map((entry) => {
      const d = new Date(entry.time);
      const hourFrac = d.getHours() + d.getMinutes() / 60;
      const x = padX + (hourFrac / 24) * (chartW - padX * 2);
      const y = padY + (1 - (moodValue[entry.mood] - 1) / 4) * (chartH - padY * 2);
      return { x, y };
    });
    const moodPolyline = moodPoints.map((p) => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');

    const habitsCheckedToday = this.habits.filter((h) => this.habitLog[todayKeyStr] && this.habitLog[todayKeyStr][h.id]).length;
    const lastMood = todayMoods.length > 0 ? MOOD_OPTIONS.find((m) => m.value === todayMoods[todayMoods.length - 1].mood) : null;
    const glanceCards = [
      { anchor: 'tasksSection', icon: '✓', label: 'Tasks', value: `${doneCount}/${s.tasks.length}`, sub: openTasks.length === 0 ? 'All done' : `${openTasks.length} left` },
      { anchor: 'habitsSection', icon: '🔥', label: 'Habits', value: `${habitsCheckedToday}/${this.habits.length}`, sub: this.habits.length === 0 ? 'None yet' : 'Checked today' },
      { anchor: 'remindersSection', icon: '📌', label: 'Reminders', value: String(remindersToday.length), sub: remindersToday.length === 0 ? 'Nothing due' : 'For today' },
      { anchor: 'moodSection', icon: lastMood ? lastMood.emoji : '❔', label: 'Mood', value: lastMood ? lastMood.label : 'Not logged', sub: 'Right now' },
    ];

    return `
      <div id="scrollArea" style="flex:1;overflow-y:auto;padding:40px 48px;box-sizing:border-box;${this._animate ? 'animation:fadeIn .3s ease' : ''}">
        <div style="margin-bottom:24px">
          <div style="font-size:13px;color:${T.textMuted};font-weight:500;margin-bottom:4px">${todayLabel}</div>
          <div style="font-family:'Manrope',sans-serif;font-weight:800;font-size:30px;letter-spacing:-0.02em">${timeGreeting}, <span style="color:${T.accentText};font-style:italic">${esc(firstName)}</span></div>
        </div>
        <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:28px">
          ${glanceCards.map((c) => `
            <div onclick="document.getElementById('${c.anchor}').scrollIntoView({behavior:'smooth',block:'start'})" style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;padding:16px;cursor:pointer">
              <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
                <div style="font-size:16px">${c.icon}</div>
                <div style="font-size:12px;color:${T.textMuted};font-weight:500">${c.label}</div>
              </div>
              <div style="font-family:'Manrope',sans-serif;font-size:20px;font-weight:800;letter-spacing:-0.01em;color:${T.text}">${esc(c.value)}</div>
              <div style="font-size:11px;color:${T.textMuted};margin-top:2px">${esc(c.sub)}</div>
            </div>`).join('')}
        </div>
        <div style="display:flex;gap:24px;align-items:flex-start;flex-wrap:wrap">
        <div style="flex:2 1 420px;min-width:0">
        <div id="tasksSection" style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;overflow:hidden;margin-bottom:24px">
          <div style="padding:18px 22px;border-bottom:1px solid ${T.divider}">
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:${s.tasks.length > 0 ? '10px' : '0'}">
              <div style="font-weight:700;font-size:15px">Today's tasks</div>
              <div style="font-size:12.5px;color:${T.textMuted}">${doneCount} of ${s.tasks.length} done</div>
            </div>
            ${s.tasks.length > 0 ? `
            <div style="height:6px;border-radius:3px;background:${T.tagBg};overflow:hidden">
              <div style="height:100%;width:${Math.round((doneCount / s.tasks.length) * 100)}%;background:${T.accentText};border-radius:3px;transition:width .3s ease"></div>
            </div>` : ''}
          </div>
          ${openTasks.map((t) => `
            <div class="row-hover" style="display:flex;align-items:center;gap:10px;padding:15px 22px;border-bottom:1px solid ${T.rowDivider};flex-wrap:wrap">
              <div onclick="App.toggleTaskDone(${t.id})" style="width:19px;height:19px;border-radius:6px;border:1.5px solid ${T.unselectedBorder};background:${T.cardBg};flex:none;cursor:pointer;display:flex;align-items:center;justify-content:center"></div>
              ${s.editingTaskId === t.id ? `
              <input id="editTaskInput-${t.id}" value="${esc(t.name)}" onkeydown="App.handleEditTaskKeyDown(event,${t.id},'editTaskInput-${t.id}')" onblur="App.saveTaskName(${t.id},'editTaskInput-${t.id}')" autofocus style="flex:1;font-size:14.5px;font-weight:500;color:${T.text};border:1.5px solid ${T.accentSolid};border-radius:7px;padding:5px 8px;font-family:inherit;outline:none;background:${T.cardBg}">` : `
              <div onclick="App.startEditTask(${t.id})" style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:14.5px;font-weight:500;color:${T.text};cursor:text">${esc(t.name)}</div>`}
              <div style="font-size:12px;color:${categoryColor(t.tag, T.dark).text};background:${categoryColor(t.tag, T.dark).bg};padding:4px 10px;border-radius:6px;font-weight:500;flex:none">${esc(t.tag)}</div>
              ${s.datePickerTaskId === t.id
                ? `<input type="date" value="${t.date || ''}" onchange="App.setTaskDate(${t.id}, this.value)" autofocus style="font-size:12px;color:${T.text};border:1px solid ${T.cardBorder};border-radius:6px;padding:4px 6px;font-family:inherit;background:${T.cardBg};flex:none">`
                : `<div onclick="App.toggleDatePicker(${t.id})" title="Set date" style="font-size:11.5px;font-weight:600;color:${t.date && t.date < todayKeyStr ? T.orangeText : (t.date ? T.accentText : T.textMuted)};background:${t.date && t.date < todayKeyStr ? T.orangeSoftBg : (t.date ? T.accentSoftBg : T.tagBg)};padding:4px 10px;border-radius:6px;cursor:pointer;flex:none;white-space:nowrap">${t.date ? (t.date < todayKeyStr ? `Overdue · ${formatTaskDate(t.date)}` : formatTaskDate(t.date)) : '+ Date'}</div>`}
              <div onclick="App.focusOnTask(${t.id})" style="font-size:12.5px;font-weight:600;color:${T.accentText};cursor:pointer;padding:6px 12px;border-radius:7px;flex:none">Focus →</div>
              <div class="remove-btn" onclick="App.removeTask(${t.id})" style="width:24px;height:24px;border-radius:7px;display:flex;align-items:center;justify-content:center;cursor:pointer;color:${T.textMuted};flex:none">
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none"><path d="M6 6l12 12M18 6L6 18" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"></path></svg>
              </div>
            </div>
            ${(t.subtasks || []).map((st) => `
              <div class="row-hover" style="display:flex;align-items:center;gap:10px;padding:11px 22px 11px 50px;border-bottom:1px solid ${T.rowDivider};flex-wrap:wrap">
                <div onclick="App.toggleSubtaskDone(${t.id},${st.id})" style="width:15px;height:15px;border-radius:5px;border:1.5px solid ${st.done ? '${T.accentSolid}' : T.unselectedBorder};background:${st.done ? '${T.accentSolid}' : T.cardBg};flex:none;cursor:pointer"></div>
                ${s.editingSubtask && s.editingSubtask.taskId === t.id && s.editingSubtask.subId === st.id ? `
                <input id="editSubtaskInput-${t.id}-${st.id}" value="${esc(st.name)}" onkeydown="App.handleEditSubtaskKeyDown(event,${t.id},${st.id},'editSubtaskInput-${t.id}-${st.id}')" onblur="App.saveSubtaskName(${t.id},${st.id},'editSubtaskInput-${t.id}-${st.id}')" autofocus style="flex:1;font-size:13px;font-weight:500;color:${T.text};border:1.5px solid ${T.accentSolid};border-radius:6px;padding:4px 7px;font-family:inherit;outline:none;background:${T.cardBg}">` : `
                <div onclick="App.startEditSubtask(${t.id},${st.id})" style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:13px;font-weight:500;text-decoration:${st.done ? 'line-through' : 'none'};color:${st.done ? T.textMuted : T.text};cursor:text">${esc(st.name)}</div>`}
                ${s.datePickerSubtask && s.datePickerSubtask.taskId === t.id && s.datePickerSubtask.subId === st.id
                  ? `<input type="date" value="${st.date || ''}" onchange="App.setSubtaskDate(${t.id},${st.id}, this.value)" autofocus style="font-size:11.5px;color:${T.text};border:1px solid ${T.cardBorder};border-radius:6px;padding:3px 5px;font-family:inherit;background:${T.cardBg};flex:none">`
                  : `<div onclick="App.toggleSubtaskDatePicker(${t.id},${st.id})" title="Set date" style="font-size:11px;font-weight:600;color:${st.date ? T.accentText : T.textMuted};background:${st.date ? T.accentSoftBg : T.tagBg};padding:3px 8px;border-radius:6px;cursor:pointer;flex:none;white-space:nowrap">${st.date ? formatTaskDate(st.date) : '+ Date'}</div>`}
                <div class="remove-btn" onclick="App.removeSubtask(${t.id},${st.id})" style="width:20px;height:20px;border-radius:6px;display:flex;align-items:center;justify-content:center;cursor:pointer;color:${T.textMuted};flex:none">
                  <svg width="11" height="11" viewBox="0 0 24 24" fill="none"><path d="M6 6l12 12M18 6L6 18" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"></path></svg>
                </div>
              </div>`).join('')}
            ${s.expandedSubtaskTaskId === t.id ? `
              <div style="display:flex;align-items:center;gap:8px;padding:10px 22px 10px 50px;border-bottom:1px solid ${T.rowDivider}">
                <input id="subtaskInput-${t.id}" onkeydown="App.handleSubtaskKeyDown(event,${t.id},'subtaskInput-${t.id}')" placeholder="Add a smaller step…" style="flex:1;border:1.5px dashed ${T.inputBorder};border-radius:8px;padding:7px 10px;font-size:13px;font-family:inherit;outline:none;color:${T.text};background:transparent">
                <div onclick="App.addSubtaskToTask(${t.id},'subtaskInput-${t.id}')" style="font-size:12px;font-weight:600;color:${T.accentText};cursor:pointer;padding:6px 12px;border-radius:7px;background:${T.accentSoftBg}">Add</div>
                <div onclick="App.toggleSubtaskInput(${t.id})" style="font-size:12px;font-weight:600;color:${T.textMuted};cursor:pointer;padding:6px 8px">Done</div>
              </div>` : `
              <div onclick="App.toggleSubtaskInput(${t.id})" style="padding:9px 22px 9px 50px;border-bottom:1px solid ${T.rowDivider};font-size:12.5px;font-weight:600;color:${T.accentText};cursor:pointer">+ Add step</div>`}`).join('')}
          <div style="padding:14px 22px">
            <div style="display:flex;align-items:center;gap:10px">
              <div style="width:19px;height:19px;border-radius:6px;border:1.5px dashed ${T.unselectedBorder};flex:none"></div>
              <input id="newTaskInput" onkeydown="App.handleTaskKeyDown(event)" placeholder="Add a task…" style="flex:1;border:none;outline:none;font-size:14.5px;font-family:inherit;font-weight:500;color:${T.text};background:transparent">
              <select id="newTaskCategory" style="border:1px solid ${T.cardBorder};border-radius:7px;padding:6px 8px;font-size:12.5px;font-family:inherit;color:${T.text};background:${T.cardBg};cursor:pointer;color-scheme:${T.dark ? 'dark' : 'light'}">
                ${s.categories.map((c) => `<option value="${c}">${c}</option>`).join('')}
              </select>
              <input type="date" id="newTaskDate" title="Due date" style="border:1px solid ${T.cardBorder};border-radius:7px;padding:6px 8px;font-size:12.5px;font-family:inherit;color:${T.text};background:${T.cardBg};cursor:pointer">
              <div onclick="App.addTask()" style="font-size:12.5px;font-weight:600;color:${T.accentText};cursor:pointer;padding:6px 14px;border-radius:7px;background:${T.accentSoftBg}">Add</div>
            </div>
            <div style="padding-left:29px;margin-top:8px;display:flex;flex-direction:column;gap:6px">
              ${s.newTaskDraftSubtasks.map((name, i) => `
                <div style="display:flex;align-items:center;gap:8px">
                  <div style="font-size:12.5px;color:${T.textMuted};flex:1">— ${esc(name)}</div>
                  <div class="remove-btn" onclick="App.removeDraftSubtask(${i})" style="width:18px;height:18px;border-radius:5px;display:flex;align-items:center;justify-content:center;cursor:pointer;color:${T.textMuted};flex:none">
                    <svg width="10" height="10" viewBox="0 0 24 24" fill="none"><path d="M6 6l12 12M18 6L6 18" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"></path></svg>
                  </div>
                </div>`).join('')}
              <div style="display:flex;align-items:center;gap:8px">
                <input id="newTaskDraftSubtaskInput" onkeydown="App.handleDraftSubtaskKeyDown(event)" placeholder="+ Add a step for this task (optional)" style="flex:1;border:1.5px dashed ${T.inputBorder};border-radius:8px;padding:7px 10px;font-size:12.5px;font-family:inherit;outline:none;color:${T.text};background:transparent">
                <div onclick="App.addDraftSubtask()" style="font-size:12px;font-weight:600;color:${T.accentText};cursor:pointer;padding:6px 12px;border-radius:7px;background:${T.accentSoftBg}">Add step</div>
              </div>
            </div>
          </div>
        </div>

        <div id="moodSection" style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;padding:20px">
          <div style="font-weight:700;font-size:15px;margin-bottom:14px">How are you feeling right now?</div>
          <div style="display:flex;align-items:center;gap:10px">
            ${MOOD_OPTIONS.map((m) => `
              <div onclick="App.logMood('${m.value}')" title="${m.label}" style="font-size:26px;width:46px;height:46px;border-radius:12px;display:flex;align-items:center;justify-content:center;cursor:pointer;background:${s.lastLoggedMood === m.value ? T.accentSoftBg : 'transparent'};border:1.5px solid ${s.lastLoggedMood === m.value ? '${T.accentSolid}' : T.cardBorder}">${m.emoji}</div>`).join('')}
          </div>
          ${todayMoods.length > 0 ? `
            <div style="margin-top:18px">
              <div style="font-size:11.5px;font-weight:600;color:${T.textMuted};margin-bottom:8px;text-transform:uppercase;letter-spacing:0.04em">Today's mood trend</div>
              <svg viewBox="0 0 ${chartW} ${chartH}" style="width:100%;height:${chartH}px;margin-bottom:10px" preserveAspectRatio="none">
                <line x1="${padX}" y1="${chartH / 2}" x2="${chartW - padX}" y2="${chartH / 2}" stroke="${T.divider}" stroke-width="1" stroke-dasharray="3,3"></line>
                ${moodPoints.length > 1 ? `<polyline points="${moodPolyline}" fill="none" stroke="oklch(0.6 0.13 260)" stroke-width="2"></polyline>` : ''}
                ${moodPoints.map((p) => `<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="3" fill="oklch(0.6 0.13 260)"></circle>`).join('')}
              </svg>
              <div style="display:flex;flex-wrap:wrap;gap:8px">
                ${todayMoods.slice().reverse().map((entry) => {
                  const opt = MOOD_OPTIONS.find((m) => m.value === entry.mood);
                  const timeLabel = new Date(entry.time).toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
                  return `<div style="font-size:12px;color:${T.textMuted};background:${T.tagBg};padding:5px 10px;border-radius:20px;display:flex;align-items:center;gap:5px">${opt ? opt.emoji : ''} ${timeLabel}</div>`;
                }).join('')}
              </div>
            </div>` : ''}
        </div>
        </div>

        <div style="flex:none;width:300px">
          <div id="habitsSection" style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;overflow:hidden;margin-bottom:24px">
            <div style="padding:18px 22px;border-bottom:1px solid ${T.divider};display:flex;align-items:center;justify-content:space-between">
              <div style="font-weight:700;font-size:15px">Habits</div>
              <div onclick="App.toggleNewHabitForm()" style="font-size:12px;font-weight:600;color:${T.accentText};cursor:pointer">${s.showNewHabitForm ? 'Cancel' : '+ New habit'}</div>
            </div>
            ${s.showNewHabitForm ? `
            <div style="padding:14px 22px;border-bottom:1px solid ${T.rowDivider};display:flex;flex-direction:column;gap:10px">
              <input id="newHabitName" placeholder="Habit name…" style="border:1.5px dashed ${T.inputBorder};border-radius:8px;padding:8px 10px;font-size:13px;font-family:inherit;outline:none;color:${T.text};background:transparent">
              <div style="display:flex;flex-wrap:wrap;gap:5px">
                ${HABIT_ICONS.map((icon) => `<div onclick="App.selectNewHabitIcon('${icon}')" style="font-size:16px;width:28px;height:28px;border-radius:7px;display:flex;align-items:center;justify-content:center;cursor:pointer;background:${s.newHabitIcon === icon ? T.accentSoftBg : 'transparent'};border:1.5px solid ${s.newHabitIcon === icon ? '${T.accentSolid}' : T.cardBorder}">${icon}</div>`).join('')}
              </div>
              <div>
                <div style="font-size:11.5px;color:${T.textMuted};margin-bottom:6px">Days per week to check in</div>
                <div style="display:flex;gap:5px">
                  ${[1, 2, 3, 4, 5, 6, 7].map((n) => `<div onclick="App.setNewHabitTarget(${n})" style="width:26px;height:26px;border-radius:7px;display:flex;align-items:center;justify-content:center;cursor:pointer;font-size:12px;font-weight:600;background:${s.newHabitTarget === n ? T.accentSoftBg : 'transparent'};border:1.5px solid ${s.newHabitTarget === n ? '${T.accentSolid}' : T.cardBorder};color:${s.newHabitTarget === n ? T.accentText : T.textMuted}">${n}</div>`).join('')}
                </div>
              </div>
              <div onclick="App.addHabit()" style="align-self:flex-start;font-size:12.5px;font-weight:600;color:${T.accentText};cursor:pointer;padding:7px 14px;border-radius:8px;background:${T.accentSoftBg}">Add habit</div>
            </div>` : ''}
            ${this.habits.length === 0 ? `<div style="padding:22px;font-size:13.5px;color:${T.textMuted}">No habits yet — tap "+ New habit" to create one.</div>` : this.habits.map((h) => {
              const checkedToday = !!(this.habitLog[todayKeyStr] && this.habitLog[todayKeyStr][h.id]);
              const streak = this.habitStreak(h.id);
              const target = h.targetPerWeek || 7;
              const doneThisWeek = this.checkInsThisWeek(h.id);
              return `
              <div style="display:flex;align-items:center;gap:12px;padding:14px 22px;border-bottom:1px solid ${T.rowDivider}">
                <div style="font-size:20px;flex:none">${h.icon}</div>
                <div style="flex:1;min-width:0">
                  <div style="font-size:13.5px;font-weight:500;color:${T.text};overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(h.name)}</div>
                  <div style="font-size:11.5px;color:${T.textMuted}">${streak} day${streak === 1 ? '' : 's'} streak · ${doneThisWeek}/${target} this week</div>
                </div>
                <div onclick="App.toggleHabitToday(${h.id})" style="width:26px;height:26px;border-radius:8px;border:1.5px solid ${checkedToday ? '${T.accentSolid}' : T.unselectedBorder};background:${checkedToday ? '${T.accentSolid}' : 'transparent'};flex:none;cursor:pointer;display:flex;align-items:center;justify-content:center">
                  ${checkedToday ? '<svg width="13" height="13" viewBox="0 0 24 24" fill="none"><path d="M5 13l4 4L19 7" stroke="white" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"></path></svg>' : ''}
                </div>
                <div class="remove-btn" onclick="App.removeHabit(${h.id})" style="width:20px;height:20px;border-radius:6px;display:flex;align-items:center;justify-content:center;cursor:pointer;color:${T.textMuted};flex:none">
                  <svg width="11" height="11" viewBox="0 0 24 24" fill="none"><path d="M6 6l12 12M18 6L6 18" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"></path></svg>
                </div>
              </div>`;
            }).join('')}
          </div>
          <div id="remindersSection" style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;overflow:hidden">
            <div style="padding:18px 22px;border-bottom:1px solid ${T.divider};font-weight:700;font-size:15px">Reminders</div>
            ${remindersToday.length === 0 ? `<div style="padding:22px;font-size:13.5px;color:${T.textMuted}">No reminders for today.</div>` : remindersToday.map((r) => `
              <div style="display:flex;align-items:center;gap:10px;padding:14px 22px;border-bottom:1px solid ${T.rowDivider}">
                ${s.editingReminderId === r.id ? `
                <input id="editTodayReminderInput-${r.id}" value="${esc(r.text)}" onkeydown="App.handleEditReminderKeyDown(event,${r.id},'editTodayReminderInput-${r.id}')" onblur="App.saveReminderText(${r.id},'editTodayReminderInput-${r.id}')" autofocus style="flex:1;font-size:13.5px;color:${T.text};border:1.5px solid ${T.accentSolid};border-radius:7px;padding:5px 8px;font-family:inherit;outline:none;background:${T.cardBg}">` : `
                <div onclick="App.startEditReminder(${r.id})" style="flex:1;font-size:13.5px;color:${T.text};cursor:text">${esc(r.text)}</div>`}
                <div class="remove-btn" onclick="App.removeReminder(${r.id})" style="width:20px;height:20px;border-radius:6px;display:flex;align-items:center;justify-content:center;cursor:pointer;color:${T.textMuted};flex:none">
                  <svg width="11" height="11" viewBox="0 0 24 24" fill="none"><path d="M6 6l12 12M18 6L6 18" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"></path></svg>
                </div>
              </div>`).join('')}
            <div style="padding:14px 22px;display:flex;flex-direction:column;gap:8px">
              <input id="newTodayReminderText" onkeydown="App.handleReminderKeyDown(event,'addTodayReminder')" placeholder="Add a reminder…" style="border:1.5px dashed ${T.inputBorder};border-radius:8px;padding:8px 10px;font-size:13px;font-family:inherit;outline:none;color:${T.text};background:transparent">
              <div onclick="App.addTodayReminder()" style="align-self:flex-start;font-size:12.5px;font-weight:600;color:${T.accentText};cursor:pointer;padding:7px 14px;border-radius:8px;background:${T.accentSoftBg}">Add</div>
            </div>
          </div>
        </div>
        </div>
      </div>`;
  },

  renderStats(T) {
    const days = this.historyForLastNDays(7);
    const weekDayLabels = days.map((d) => d.date.toLocaleDateString(undefined, { weekday: 'short' }));
    const weekMinutes = days.map((d) => Math.round(d.focusSeconds / 60));
    const maxWeekMin = Math.max(1, ...weekMinutes);
    const totals = this.allTimeTotals();
    const totalMinutes = Math.round(totals.focusSeconds / 60);
    const totalHours = Math.floor(totalMinutes / 60);
    const totalFocusLabel = totalHours > 0 ? `${totalHours}h ${totalMinutes % 60}m` : `${totalMinutes}m`;
    const overallStats = [
      { label: 'Total focus time', value: totalFocusLabel, color: T.text },
      { label: 'Sessions completed', value: String(totals.sessions), color: T.accentText },
      { label: 'Longest streak', value: `${this.longestStreak()} days`, color: T.orangeText },
    ];
    return `
      <div id="scrollArea" style="flex:1;overflow-y:auto;padding:40px 48px;box-sizing:border-box;${this._animate ? 'animation:fadeIn .3s ease' : ''}">
        <div style="margin-bottom:32px">
          <div style="font-size:13px;color:${T.textMuted};font-weight:500;margin-bottom:4px">Insights</div>
          <div style="font-family:'Manrope',sans-serif;font-weight:800;font-size:28px;letter-spacing:-0.02em">Your focus stats</div>
        </div>
        <div style="display:flex;gap:24px;align-items:flex-start;flex-wrap:wrap">
          <div style="flex:2 1 420px;min-width:0">
            <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:20px">
              ${overallStats.map((stat) => `
                <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;padding:20px">
                  <div style="font-size:12.5px;color:${T.textMuted};font-weight:500;margin-bottom:8px">${stat.label}</div>
                  <div style="font-family:'Manrope',sans-serif;font-size:26px;font-weight:800;letter-spacing:-0.01em;color:${stat.color}">${stat.value}</div>
                </div>`).join('')}
            </div>
            <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;padding:24px;margin-bottom:20px">
              <div style="font-weight:700;font-size:15px;margin-bottom:20px">This week</div>
              <div style="display:flex;align-items:flex-end;gap:14px;height:140px">
                ${weekDayLabels.map((label, i) => `
                  <div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;height:100%;gap:8px">
                    <div style="font-size:11px;color:${T.textMuted};font-weight:600">${weekMinutes[i]}m</div>
                    <div style="width:100%;max-width:32px;height:${weekMinutes[i] > 0 ? Math.max(10, (weekMinutes[i] / maxWeekMin) * 100) : 3}%;border-radius:6px;background:${i === days.length - 1 ? '${T.accentSolid}' : 'oklch(0.85 0.04 152)'}"></div>
                    <div style="font-size:11.5px;color:${T.textMuted};font-weight:600">${label}</div>
                  </div>`).join('')}
              </div>
            </div>
            ${this.habits.length > 0 ? `
            <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;padding:24px;margin-bottom:20px">
              <div style="font-weight:700;font-size:15px;margin-bottom:16px">Habit streaks</div>
              <div style="display:flex;flex-direction:column;gap:12px">
                ${this.habits.map((h) => `
                  <div style="display:flex;align-items:center;gap:12px">
                    <div style="font-size:18px;flex:none">${h.icon}</div>
                    <div style="flex:1;font-size:13.5px;font-weight:500;color:${T.text}">${esc(h.name)}</div>
                    <div style="font-size:13px;font-weight:700;color:${T.accentText}">${this.habitStreak(h.id)} day${this.habitStreak(h.id) === 1 ? '' : 's'}</div>
                  </div>`).join('')}
              </div>
            </div>` : ''}
            ${this.renderDayDetail(T)}
          </div>
          <div style="flex:none;width:280px">
            ${this.renderMonthCalendar(T, true)}
          </div>
        </div>
      </div>`;
  },

  renderMonthCalendar(T, compact) {
    const s = this.state;
    const { year, month } = s.statsViewMonth;
    const firstOfMonth = new Date(year, month, 1);
    const startWeekday = firstOfMonth.getDay();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const monthLabel = firstOfMonth.toLocaleDateString(undefined, { month: 'long', year: 'numeric' });
    const todayKeyStr = dayKey(new Date());
    const weekdayLabels = ['S', 'M', 'T', 'W', 'T', 'F', 'S'];
    const intensity = (minutes) => {
      if (!minutes) return T.divider;
      if (minutes < 15) return 'oklch(0.88 0.07 152)';
      if (minutes < 30) return 'oklch(0.75 0.11 152)';
      if (minutes < 60) return 'oklch(0.62 0.14 152)';
      return 'oklch(0.5 0.15 152)';
    };
    const cells = [];
    for (let i = 0; i < startWeekday; i++) cells.push(null);
    for (let d = 1; d <= daysInMonth; d++) cells.push(d);
    const pad = compact ? 16 : 24;
    const gap = compact ? 4 : 6;
    const dayFont = compact ? '10.5px' : '12px';
    const headerFont = compact ? '13px' : '15px';
    return `
      <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;padding:${pad}px">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:${compact ? 12 : 18}px">
          <div onclick="App.goToPrevMonth()" style="cursor:pointer;padding:4px 8px;font-size:${compact ? 13 : 15}px;color:${T.textMuted}">←</div>
          <div style="font-weight:700;font-size:${headerFont}">${monthLabel}</div>
          <div onclick="App.goToNextMonth()" style="cursor:pointer;padding:4px 8px;font-size:${compact ? 13 : 15}px;color:${T.textMuted}">→</div>
        </div>
        <div style="display:grid;grid-template-columns:repeat(7,1fr);gap:${gap}px;margin-bottom:${gap}px">
          ${weekdayLabels.map((w) => `<div style="text-align:center;font-size:${compact ? 9.5 : 11}px;color:${T.textMuted};font-weight:600">${w}</div>`).join('')}
        </div>
        <div style="display:grid;grid-template-columns:repeat(7,1fr);gap:${gap}px">
          ${cells.map((d) => {
            if (!d) return '<div></div>';
            const key = `${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
            const rec = this.history[key] || { focusSeconds: 0 };
            const minutes = Math.round(rec.focusSeconds / 60);
            const isToday = key === todayKeyStr;
            const isSelected = key === s.statsSelectedDate;
            const border = isSelected ? '2px solid oklch(0.4 0.15 152)' : isToday ? `1.5px solid ${T.textMuted}` : '1px solid transparent';
            return `<div onclick="App.selectStatsDate('${key}')" style="aspect-ratio:1;border-radius:${compact ? 6 : 8}px;display:flex;align-items:center;justify-content:center;cursor:pointer;background:${intensity(minutes)};box-sizing:border-box;border:${border};font-size:${dayFont};font-weight:600;color:${minutes >= 30 ? 'white' : T.text}">${d}</div>`;
          }).join('')}
        </div>
      </div>`;
  },

  renderDayDetail(T) {
    const s = this.state;
    if (!s.statsSelectedDate) return '';
    const key = s.statsSelectedDate;
    const [y, m, d] = key.split('-').map(Number);
    const label = new Date(y, m - 1, d).toLocaleDateString(undefined, { weekday: 'long', month: 'long', day: 'numeric' });
    const rec = this.history[key] || { focusSeconds: 0, sessionsCompleted: 0 };
    const minutes = Math.round(rec.focusSeconds / 60);
    const tasksThatDay = this.taskHistory[key] || [];
    const moodsThatDay = (this.moodLog[key] || []).slice().sort((a, b) => a.time - b.time);

    const notFinishedThatDay = [];
    s.tasks.forEach((t) => {
      if (t.date === key && !t.done) notFinishedThatDay.push({ name: t.name, tag: t.tag });
      (t.subtasks || []).forEach((st) => {
        if (st.date === key && !st.done) notFinishedThatDay.push({ name: st.name, tag: t.tag });
      });
    });

    const moodValue = { awful: 1, bad: 2, okay: 3, good: 4, great: 5 };
    const chartW = 300, chartH = 90, padX = 8, padY = 10;
    const points = moodsThatDay.map((entry) => {
      const d = new Date(entry.time);
      const hourFrac = d.getHours() + d.getMinutes() / 60;
      const x = padX + (hourFrac / 24) * (chartW - padX * 2);
      const y = padY + (1 - (moodValue[entry.mood] - 1) / 4) * (chartH - padY * 2);
      return { x, y, entry };
    });
    const polylinePoints = points.map((p) => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');

    return `
      <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;padding:24px">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px">
          <div style="font-weight:700;font-size:15px">${label}</div>
          <div onclick="App.selectStatsDate(null)" style="cursor:pointer;font-size:12.5px;color:${T.textMuted}">Close ✕</div>
        </div>
        <div style="display:flex;gap:28px;margin-bottom:20px">
          <div>
            <div style="font-size:12px;color:${T.textMuted};margin-bottom:4px">Focused</div>
            <div style="font-family:'Manrope',sans-serif;font-weight:800;font-size:20px">${minutes}m</div>
          </div>
          <div>
            <div style="font-size:12px;color:${T.textMuted};margin-bottom:4px">Sessions</div>
            <div style="font-family:'Manrope',sans-serif;font-weight:800;font-size:20px">${rec.sessionsCompleted}</div>
          </div>
        </div>
        <div style="font-size:12.5px;font-weight:600;color:${T.textMuted};margin-bottom:10px;text-transform:uppercase;letter-spacing:0.04em">Tasks finished</div>
        ${tasksThatDay.length === 0 ? `<div style="font-size:13.5px;color:${T.textMuted};margin-bottom:20px">No tasks logged as completed this day.</div>` : `<div style="margin-bottom:20px">${tasksThatDay.map((t) => `
          <div style="display:flex;align-items:center;gap:12px;padding:10px 0;border-bottom:1px solid ${T.rowDivider}">
            <div style="flex:1;font-size:13.5px;font-weight:500">${esc(t.name)}</div>
            <div style="font-size:11.5px;color:${categoryColor(t.tag, T.dark).text};background:${categoryColor(t.tag, T.dark).bg};padding:3px 8px;border-radius:6px;font-weight:500">${esc(t.tag)}</div>
          </div>`).join('')}</div>`}
        <div style="font-size:12.5px;font-weight:600;color:${T.textMuted};margin-bottom:10px;text-transform:uppercase;letter-spacing:0.04em">Tasks not finished</div>
        ${notFinishedThatDay.length === 0 ? `<div style="font-size:13.5px;color:${T.textMuted};margin-bottom:20px">Nothing left unfinished for this day.</div>` : `<div style="margin-bottom:20px">${notFinishedThatDay.map((t) => `
          <div style="display:flex;align-items:center;gap:12px;padding:10px 0;border-bottom:1px solid ${T.rowDivider}">
            <div style="flex:1;font-size:13.5px;font-weight:500;color:${T.text}">${esc(t.name)}</div>
            <div style="font-size:11.5px;color:${T.orangeText};background:${T.orangeSoftBg};padding:3px 8px;border-radius:6px;font-weight:500">${esc(t.tag)}</div>
          </div>`).join('')}</div>`}
        ${this.habits.length > 0 ? `
        <div style="font-size:12.5px;font-weight:600;color:${T.textMuted};margin-bottom:10px;text-transform:uppercase;letter-spacing:0.04em">Habits checked</div>
        <div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:20px">
          ${this.habits.map((h) => {
            const checked = !!(this.habitLog[key] && this.habitLog[key][h.id]);
            return `<div style="font-size:12.5px;color:${checked ? T.text : T.textMuted};background:${checked ? T.accentSoftBg : T.tagBg};border:1px solid ${checked ? '${T.accentSolid}' : 'transparent'};padding:5px 10px;border-radius:20px;display:flex;align-items:center;gap:5px">${h.icon} ${esc(h.name)} ${checked ? '✓' : '—'}</div>`;
          }).join('')}
        </div>` : ''}
        <div style="font-size:12.5px;font-weight:600;color:${T.textMuted};margin-bottom:10px;text-transform:uppercase;letter-spacing:0.04em">Mood trend</div>
        ${moodsThatDay.length === 0 ? `<div style="font-size:13.5px;color:${T.textMuted}">No mood check-ins logged this day.</div>` : `
          <svg viewBox="0 0 ${chartW} ${chartH}" style="width:100%;height:${chartH}px;margin-bottom:10px" preserveAspectRatio="none">
            <line x1="${padX}" y1="${chartH / 2}" x2="${chartW - padX}" y2="${chartH / 2}" stroke="${T.divider}" stroke-width="1" stroke-dasharray="3,3"></line>
            ${points.length > 1 ? `<polyline points="${polylinePoints}" fill="none" stroke="oklch(0.6 0.13 260)" stroke-width="2"></polyline>` : ''}
            ${points.map((p) => `<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="3" fill="oklch(0.6 0.13 260)"></circle>`).join('')}
          </svg>
          <div style="display:flex;flex-wrap:wrap;gap:8px">
            ${moodsThatDay.map((entry) => {
              const opt = MOOD_OPTIONS.find((m) => m.value === entry.mood);
              const timeLabel = new Date(entry.time).toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
              return `<div style="font-size:12.5px;color:${T.text};background:${T.tagBg};padding:5px 10px;border-radius:20px;display:flex;align-items:center;gap:5px">${opt ? opt.emoji : ''} ${timeLabel}</div>`;
            }).join('')}
          </div>`}
      </div>`;
  },

  renderInbox(T) {
    const s = this.state;
    const showGenerateButton = s.aiStage !== 'ready';
    const generateButtonLabel = s.aiStage === 'generating' ? 'Generating…' : 'Generate summary & to-dos';
    const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const last24hCutoff = new Date(Date.now() - 24 * 60 * 60 * 1000);
    const recentEmails = s.emails.filter((m) => m.receivedAt >= last24hCutoff);
    const earlierThisWeekEmails = s.emails.filter((m) => m.receivedAt < last24hCutoff);
    const renderEmailRow = (mail) => {
      const expanded = s.expandedEmailId === mail.id;
      return `
            <div class="row-hover" onclick="App.toggleEmailExpanded('${mail.id}')" style="display:flex;align-items:flex-start;gap:14px;padding:16px 22px;border-bottom:1px solid ${T.rowDivider};cursor:pointer">
              <div style="width:34px;height:34px;border-radius:50%;background:${mail.avatarColor};flex:none;display:flex;align-items:center;justify-content:center;color:white;font-weight:700;font-size:12.5px">${esc(mail.initial)}</div>
              <div style="flex:1;min-width:0">
                <div style="display:flex;justify-content:space-between;gap:12px;align-items:center">
                  <div style="display:flex;align-items:center;gap:7px;min-width:0">
                    ${!mail.isRead ? `<div title="Unread" style="width:8px;height:8px;border-radius:50%;background:oklch(0.6 0.18 260);flex:none"></div>` : ''}
                    <div style="font-size:14px;font-weight:${mail.isRead ? '600' : '800'};min-width:0;overflow:hidden;white-space:nowrap;text-overflow:ellipsis">${esc(mail.sender)}</div>
                    ${this.emailNotes[mail.id] ? `<div title="Has a note" style="width:7px;height:7px;border-radius:50%;background:${T.orangeText};flex:none"></div>` : ''}
                  </div>
                  <div style="font-size:12px;color:${T.textMuted};flex:none">${esc(mail.time)}</div>
                </div>
                <div style="font-size:13.5px;font-weight:${mail.isRead ? '500' : '700'};color:${T.text};margin:2px 0">${esc(mail.subject)}</div>
                <div style="font-size:12.5px;color:${T.textMuted};white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${esc(mail.snippet)}</div>
              </div>
            </div>
            ${expanded ? `
            <div onclick="event.stopPropagation()" style="display:flex;gap:20px;padding:18px 22px;border-bottom:1px solid ${T.rowDivider};background:${T.subtleBg}">
              <div style="flex:2;min-width:0;font-size:13.5px;line-height:1.6;color:${T.text};white-space:pre-wrap;word-wrap:break-word;max-height:340px;overflow-y:auto">${esc(mail.body || mail.snippet)}</div>
              <div style="flex:1;min-width:180px">
                <div style="font-size:11.5px;font-weight:600;color:${T.textMuted};text-transform:uppercase;letter-spacing:0.04em;margin-bottom:8px">Your notes</div>
                <textarea onchange="App.setEmailNote('${mail.id}', this.value)" placeholder="Add a note about this email…" style="width:100%;min-height:140px;border:1.5px solid ${T.cardBorder};border-radius:10px;padding:10px;font-size:13px;font-family:inherit;color:${T.text};background:${T.cardBg};resize:vertical;box-sizing:border-box">${esc(this.emailNotes[mail.id] || '')}</textarea>
              </div>
            </div>` : ''}`;
    };
    return `
      <div id="scrollArea" style="flex:1;overflow-y:auto;padding:40px 48px;box-sizing:border-box;${this._animate ? 'animation:fadeIn .3s ease' : ''}">
        <div style="display:flex;align-items:flex-end;justify-content:space-between;margin-bottom:28px">
          <div>
            <div style="font-size:13px;color:${T.textMuted};font-weight:500;margin-bottom:4px">Outlook inbox</div>
            <div style="font-family:'Manrope',sans-serif;font-weight:800;font-size:28px;letter-spacing:-0.02em">Mail to tasks</div>
          </div>
          <div style="display:flex;align-items:center;gap:10px">
            ${showGenerateButton ? `
              <div onclick="${s.aiStage === 'generating' ? '' : 'App.generateSummary()'}" style="display:flex;align-items:center;gap:9px;padding:12px 20px;background:${T.accentSolid};color:white;border-radius:11px;font-weight:600;font-size:14px;cursor:pointer;box-shadow:0 4px 14px ${T.dark ? 'oklch(0.52 0.15 152 / 0.32)' : 'oklch(0.36 0.09 152 / 0.28)'}">
                <div style="width:14px;height:14px;border-radius:50%;background:white;opacity:0.85"></div>
                ${generateButtonLabel}
              </div>` : ''}
          </div>
        </div>

        ${s.emailFetchError ? `<div style="background:${T.navySoftBg};border:1px solid ${T.navySoftBorder};border-radius:12px;padding:16px 20px;margin-bottom:24px;font-size:13.5px;color:${T.navyText}">${esc(s.emailFetchError)}</div>` : ''}

        ${s.aiStage === 'generating' ? `
          <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;padding:24px;margin-bottom:24px;display:flex;align-items:center;gap:14px">
            <div style="width:22px;height:22px;border-radius:50%;border:2.5px solid ${T.spinnerTrack};border-top-color:${T.accentSolid};animation:spin 0.8s linear infinite"></div>
            <div style="font-size:14px;font-weight:500;color:${T.text}">Asking Claude to read your inbox and draft a summary…</div>
          </div>` : ''}

        ${s.aiStage === 'error' ? `
          <div style="background:${T.errorSoftBg};border:1px solid ${T.errorSoftBorder};border-radius:14px;padding:20px;margin-bottom:24px;color:${T.errorText};font-size:14px">${esc(s.aiError)}</div>` : ''}

        ${s.aiStage === 'ready' ? `
          <div style="background:${T.navySoftBg};border:1px solid ${T.navySoftBorder};border-radius:14px;padding:24px;margin-bottom:24px">
            <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px">
              <div style="width:18px;height:18px;border-radius:5px;background:${T.navyIcon}"></div>
              <div style="font-weight:700;font-size:14.5px;color:${T.navyText};letter-spacing:0.01em">AI summary (Claude)</div>
            </div>
            <div style="font-size:14px;line-height:1.6;color:${T.text}">${esc(s.aiSummaryText)}</div>
          </div>
          <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;overflow:hidden;margin-bottom:24px">
            <div style="padding:18px 22px;border-bottom:1px solid ${T.divider};font-weight:700;font-size:15px">Suggested to-dos from your mail</div>
            ${s.suggestedTodos.map((todo) => `
              <div style="display:flex;align-items:center;gap:14px;padding:15px 22px;border-bottom:1px solid ${T.rowDivider}">
                <div style="flex:1">
                  <div style="font-size:14.5px;font-weight:500;margin-bottom:3px">${esc(todo.name)}</div>
                  <div style="font-size:12px;color:${T.textMuted}">From: ${esc(todo.source)}${todo.date ? ` · Due ${formatTaskDate(todo.date)}` : ''}</div>
                </div>
                <div onclick="App.addTodoToTasks(${todo.id})" style="font-size:12.5px;font-weight:600;color:${todo.added ? T.textMuted : T.accentText};cursor:pointer;padding:7px 14px;border-radius:8px;border:1.5px solid ${todo.added ? T.unselectedBorder : T.accentSolid};background:${todo.added ? T.rowDivider : T.accentSoftBg}">${todo.added ? 'Added' : 'Add to tasks'}</div>
              </div>`).join('')}
          </div>` : ''}

        <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;overflow:hidden;margin-bottom:24px">
          <div style="padding:18px 22px;border-bottom:1px solid ${T.divider};font-weight:700;font-size:15px">Last 24 hours</div>
          ${recentEmails.length === 0 ? `<div style="padding:22px;font-size:13.5px;color:${T.textMuted}">No emails in the last 24 hours.</div>` : recentEmails.map(renderEmailRow).join('')}
        </div>

        <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;overflow:hidden">
          <div style="padding:18px 22px;border-bottom:1px solid ${T.divider};font-weight:700;font-size:15px">Earlier this week (${dayNames[0]}–${dayNames[new Date().getDay()]})</div>
          ${earlierThisWeekEmails.length === 0 ? `<div style="padding:22px;font-size:13.5px;color:${T.textMuted}">No earlier emails this week.</div>` : earlierThisWeekEmails.map(renderEmailRow).join('')}
        </div>
      </div>`;
  },

  itemsForDate(dateStr) {
    let todoCount = 0;
    this.state.tasks.forEach((t) => {
      if (t.date === dateStr) todoCount++;
      (t.subtasks || []).forEach((st) => { if (st.date === dateStr) todoCount++; });
    });
    const eventCount = this.state.events.filter((e) => this.eventOccursOnDate(e, dateStr)).length;
    return { todoCount, eventCount, total: todoCount + eventCount };
  },

  renderCalendar(T) {
    const s = this.state;
    const scanButtonLabel = s.eventScanStage === 'scanning' ? 'Scanning…' : 'Scan emails for events';
    const modes = [
      { key: 'day', label: 'Day' },
      { key: 'month', label: 'Month' },
      { key: 'year', label: 'Year' },
    ];
    return `
      <div id="scrollArea" style="flex:1;overflow-y:auto;padding:40px 48px;box-sizing:border-box;${this._animate ? 'animation:fadeIn .3s ease' : ''}">
        <div style="display:flex;align-items:flex-end;justify-content:space-between;margin-bottom:28px">
          <div>
            <div style="font-size:13px;color:${T.textMuted};font-weight:500;margin-bottom:4px">Calendar</div>
            <div style="font-family:'Manrope',sans-serif;font-weight:800;font-size:28px;letter-spacing:-0.02em">To-dos, events & important dates</div>
          </div>
          <div onclick="${s.eventScanStage === 'scanning' ? '' : 'App.scanEmailsForEvents()'}" style="display:flex;align-items:center;gap:9px;padding:12px 20px;background:${T.cardBg};border:1.5px solid ${T.cardBorder};color:${T.text};border-radius:11px;font-weight:600;font-size:14px;cursor:pointer">
            ${scanButtonLabel}
          </div>
        </div>

        ${s.eventScanStage === 'error' ? `<div style="background:${T.errorSoftBg};border:1px solid ${T.errorSoftBorder};border-radius:12px;padding:16px 20px;margin-bottom:24px;font-size:13.5px;color:${T.errorText}">${esc(s.eventScanError)}</div>` : ''}

        ${s.proposedEvents.length > 0 ? `
          <div style="background:${T.navySoftBg};border:1px solid ${T.navySoftBorder};border-radius:14px;overflow:hidden;margin-bottom:24px">
            <div style="padding:18px 22px;border-bottom:1px solid ${T.divider};font-weight:700;font-size:15px;color:${T.navyText}">Events found in your mail — review before adding</div>
            ${s.proposedEvents.map((e) => `
              <div style="display:flex;align-items:center;gap:14px;padding:15px 22px;border-bottom:1px solid ${T.rowDivider}">
                <div style="flex:1">
                  <div style="font-size:14.5px;font-weight:500;margin-bottom:3px">${esc(e.name)}</div>
                  <div style="font-size:12px;color:${T.textMuted}">${formatTaskDate(e.date)} · From: ${esc(e.source)}</div>
                </div>
                <div onclick="App.dismissProposedEvent(${e.id})" style="font-size:12.5px;font-weight:600;color:${T.textMuted};cursor:pointer;padding:7px 14px;border-radius:8px">Dismiss</div>
                <div onclick="App.acceptProposedEvent(${e.id})" style="font-size:12.5px;font-weight:600;color:${T.accentText};cursor:pointer;padding:7px 14px;border-radius:8px;border:1.5px solid ${T.accentSolid};background:${T.accentSoftBg}">Add to calendar</div>
              </div>`).join('')}
          </div>` : ''}

        <div style="display:flex;gap:6px;margin-bottom:24px;background:${T.tagBg};border-radius:10px;padding:4px;width:fit-content">
          ${modes.map((m) => `
            <div onclick="App.setCalendarViewMode('${m.key}')" style="padding:7px 18px;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;background:${s.calendarViewMode === m.key ? T.cardBg : 'transparent'};color:${s.calendarViewMode === m.key ? T.text : T.textMuted};box-shadow:${s.calendarViewMode === m.key ? '0 1px 3px rgba(0,0,0,0.08)' : 'none'}">${m.label}</div>`).join('')}
        </div>

        ${s.calendarViewMode === 'day' ? this.renderCalendarDayView(T) : ''}
        ${s.calendarViewMode === 'month' ? this.renderCalendarMonthView(T) : ''}
        ${s.calendarViewMode === 'year' ? this.renderCalendarYearView(T) : ''}
      </div>`;
  },

  renderCalendarDayView(T) {
    const s = this.state;
    const selectedDate = s.calendarSelectedDate;
    const dateLabel = parseLocalDate(selectedDate).toLocaleDateString(undefined, { weekday: 'long', month: 'long', day: 'numeric' });
    const isToday = selectedDate === dayKey(new Date());

    const todosDue = [];
    s.tasks.forEach((t) => {
      if (t.date === selectedDate) todosDue.push({ kind: 'task', taskId: t.id, name: t.name, done: t.done, overdue: false });
      (t.subtasks || []).forEach((st) => {
        if (st.date === selectedDate) todosDue.push({ kind: 'subtask', taskId: t.id, subId: st.id, name: st.name, done: st.done, overdue: false });
      });
    });
    if (isToday) {
      // Carry forward anything overdue and still not done — shown here in addition to its original date.
      s.tasks.forEach((t) => {
        if (t.date && t.date < selectedDate && !t.done) todosDue.push({ kind: 'task', taskId: t.id, name: t.name, done: t.done, overdue: true });
        (t.subtasks || []).forEach((st) => {
          if (st.date && st.date < selectedDate && !st.done) todosDue.push({ kind: 'subtask', taskId: t.id, subId: st.id, name: st.name, done: st.done, overdue: true });
        });
      });
    }

    const eventsThatDay = s.events.filter((e) => this.eventOccursOnDate(e, selectedDate)).sort((a, b) => (a.time || '').localeCompare(b.time || ''));
    const remindersThatDay = s.reminders.filter((r) => r.date === selectedDate);

    return `
        <div style="display:flex;align-items:center;gap:14px;margin-bottom:24px">
          <div onclick="App.shiftCalendarDate(-1)" style="width:36px;height:36px;border-radius:9px;border:1.5px solid ${T.unselectedBorder};display:flex;align-items:center;justify-content:center;cursor:pointer;font-weight:700;font-size:16px">‹</div>
          <div style="flex:1;text-align:center;font-weight:700;font-size:15.5px">${dateLabel}${isToday ? ` <span style="font-weight:600;color:${T.accentText};font-size:12.5px">· Today</span>` : ''}</div>
          <div onclick="App.shiftCalendarDate(1)" style="width:36px;height:36px;border-radius:9px;border:1.5px solid ${T.unselectedBorder};display:flex;align-items:center;justify-content:center;cursor:pointer;font-weight:700;font-size:16px">›</div>
        </div>
        ${!isToday ? `<div onclick="App.selectCalendarDate('${dayKey(new Date())}')" style="text-align:center;margin-top:-14px;margin-bottom:24px;font-size:12.5px;font-weight:600;color:${T.accentText};cursor:pointer">Jump to today</div>` : ''}

        <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;overflow:hidden;margin-bottom:24px">
          <div style="padding:18px 22px;border-bottom:1px solid ${T.divider};font-weight:700;font-size:15px">To-dos due this day</div>
          ${todosDue.length === 0 ? `<div style="padding:22px;font-size:13.5px;color:${T.textMuted}">No to-dos scheduled for this day.</div>` : todosDue.map((td) => `
            <div style="display:flex;align-items:center;gap:14px;padding:14px 22px;border-bottom:1px solid ${T.rowDivider}">
              <div onclick="${td.kind === 'task' ? `App.toggleTaskDone(${td.taskId})` : `App.toggleSubtaskDone(${td.taskId},${td.subId})`}" style="width:17px;height:17px;border-radius:5px;border:1.5px solid ${td.done ? '${T.accentSolid}' : T.unselectedBorder};background:${td.done ? '${T.accentSolid}' : T.cardBg};flex:none;cursor:pointer"></div>
              <div style="flex:1;font-size:14px;font-weight:500;text-decoration:${td.done ? 'line-through' : 'none'};color:${td.done ? T.textMuted : T.text}">${esc(td.name)}</div>
              <div style="font-size:11.5px;color:${td.overdue ? T.orangeText : T.textMuted};background:${td.overdue ? T.orangeSoftBg : T.tagBg};padding:3px 8px;border-radius:6px;font-weight:500">${td.overdue ? 'Overdue' : (td.kind === 'task' ? 'Task' : 'Step')}</div>
            </div>`).join('')}
        </div>

        <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;overflow:hidden">
          <div style="padding:18px 22px;border-bottom:1px solid ${T.divider};font-weight:700;font-size:15px">Events & important dates</div>
          ${eventsThatDay.length === 0 ? `<div style="padding:22px;font-size:13.5px;color:${T.textMuted}">No events for this day.</div>` : eventsThatDay.map((e) => `
            <div style="display:flex;align-items:center;gap:14px;padding:14px 22px;border-bottom:1px solid ${T.rowDivider}">
              ${e.important ? `<div title="Important" style="width:8px;height:8px;border-radius:50%;background:${T.orangeText};flex:none"></div>` : ''}
              <div style="flex:1">
                <div style="font-size:14px;font-weight:500;color:${T.text}">${esc(e.name)}</div>
                ${e.endDate ? `<div style="font-size:11.5px;color:${T.accentText};margin-top:2px">${formatTaskDate(e.date)} – ${formatTaskDate(e.endDate)}</div>` : ''}
                ${e.source === 'ai' ? `<div style="font-size:11.5px;color:${T.textMuted};margin-top:2px">Added from mail</div>` : ''}
              </div>
              ${e.time ? `<div style="font-size:12px;color:${T.textMuted}">${esc(e.time)}</div>` : ''}
              <div class="remove-btn" onclick="App.removeEvent(${e.id})" style="width:22px;height:22px;border-radius:6px;display:flex;align-items:center;justify-content:center;cursor:pointer;color:${T.textMuted};flex:none">
                <svg width="12" height="12" viewBox="0 0 24 24" fill="none"><path d="M6 6l12 12M18 6L6 18" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"></path></svg>
              </div>
            </div>`).join('')}
          <div style="padding:14px 22px;display:flex;align-items:center;gap:8px;flex-wrap:wrap">
            <input id="newEventName" placeholder="Add an event or important date…" style="flex:1;min-width:160px;border:1.5px dashed ${T.inputBorder};border-radius:8px;padding:8px 10px;font-size:13.5px;font-family:inherit;outline:none;color:${T.text};background:transparent">
            <input type="date" id="newEventDate" value="${selectedDate}" title="Start date" style="border:1px solid ${T.cardBorder};border-radius:8px;padding:7px 8px;font-size:12.5px;font-family:inherit;color:${T.text};background:${T.cardBg}">
            <span style="font-size:12px;color:${T.textMuted}">to</span>
            <input type="date" id="newEventEndDate" title="End date (optional, for multi-day events)" style="border:1px solid ${T.cardBorder};border-radius:8px;padding:7px 8px;font-size:12.5px;font-family:inherit;color:${T.text};background:${T.cardBg}">
            <input type="time" id="newEventTime" style="border:1px solid ${T.cardBorder};border-radius:8px;padding:7px 8px;font-size:12.5px;font-family:inherit;color:${T.text};background:${T.cardBg}">
            <label style="display:flex;align-items:center;gap:5px;font-size:12px;color:${T.textMuted};cursor:pointer">
              <input type="checkbox" id="newEventImportant"> Important
            </label>
            <div onclick="App.addEvent()" style="font-size:12.5px;font-weight:600;color:${T.accentText};cursor:pointer;padding:8px 16px;border-radius:8px;background:${T.accentSoftBg}">Add</div>
          </div>
        </div>

        <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;overflow:hidden;margin-top:24px">
          <div style="padding:18px 22px;border-bottom:1px solid ${T.divider};font-weight:700;font-size:15px">Reminders</div>
          ${remindersThatDay.length === 0 ? `<div style="padding:22px;font-size:13.5px;color:${T.textMuted}">No reminders for this day.</div>` : remindersThatDay.map((r) => `
            <div style="display:flex;align-items:center;gap:14px;padding:14px 22px;border-bottom:1px solid ${T.rowDivider}">
              ${s.editingReminderId === r.id ? `
              <input id="editReminderInput-${r.id}" value="${esc(r.text)}" onkeydown="App.handleEditReminderKeyDown(event,${r.id},'editReminderInput-${r.id}')" onblur="App.saveReminderText(${r.id},'editReminderInput-${r.id}')" autofocus style="flex:1;font-size:14px;color:${T.text};border:1.5px solid ${T.accentSolid};border-radius:7px;padding:5px 8px;font-family:inherit;outline:none;background:${T.cardBg}">` : `
              <div onclick="App.startEditReminder(${r.id})" style="flex:1;font-size:14px;color:${T.text};cursor:text">${esc(r.text)}</div>`}
              <div class="remove-btn" onclick="App.removeReminder(${r.id})" style="width:22px;height:22px;border-radius:6px;display:flex;align-items:center;justify-content:center;cursor:pointer;color:${T.textMuted};flex:none">
                <svg width="12" height="12" viewBox="0 0 24 24" fill="none"><path d="M6 6l12 12M18 6L6 18" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"></path></svg>
              </div>
            </div>`).join('')}
          <div style="padding:14px 22px;display:flex;align-items:center;gap:8px">
            <input id="newReminderText" onkeydown="App.handleReminderKeyDown(event,'addReminderForSelectedDate')" placeholder="Add a reminder for this day…" style="flex:1;border:1.5px dashed ${T.inputBorder};border-radius:8px;padding:8px 10px;font-size:13.5px;font-family:inherit;outline:none;color:${T.text};background:transparent">
            <div onclick="App.addReminderForSelectedDate()" style="font-size:12.5px;font-weight:600;color:${T.accentText};cursor:pointer;padding:8px 16px;border-radius:8px;background:${T.accentSoftBg}">Add</div>
          </div>
        </div>`;
  },

  renderCalendarMonthView(T) {
    const s = this.state;
    const { year, month } = s.calendarViewMonth;
    const firstOfMonth = new Date(year, month, 1);
    const startWeekday = firstOfMonth.getDay();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const monthLabel = firstOfMonth.toLocaleDateString(undefined, { month: 'long', year: 'numeric' });
    const todayKeyStr = dayKey(new Date());
    const weekdayLabels = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const cells = [];
    for (let i = 0; i < startWeekday; i++) cells.push(null);
    for (let d = 1; d <= daysInMonth; d++) cells.push(d);

    const monthStartKey = `${year}-${String(month + 1).padStart(2, '0')}-01`;
    const monthEndKey = `${year}-${String(month + 1).padStart(2, '0')}-${String(daysInMonth).padStart(2, '0')}`;
    const eventsThisMonth = s.events
      .filter((e) => e.date <= monthEndKey && (e.endDate || e.date) >= monthStartKey)
      .sort((a, b) => a.date.localeCompare(b.date));

    return `
        <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;padding:24px;margin-bottom:24px">
          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px">
            <div onclick="App.calendarShiftMonth(-1)" style="cursor:pointer;padding:4px 12px;font-size:15px;color:${T.textMuted}">←</div>
            <div style="font-weight:700;font-size:16px">${monthLabel}</div>
            <div onclick="App.calendarShiftMonth(1)" style="cursor:pointer;padding:4px 12px;font-size:15px;color:${T.textMuted}">→</div>
          </div>
          <div style="display:grid;grid-template-columns:repeat(7,1fr);gap:8px;margin-bottom:8px">
            ${weekdayLabels.map((w) => `<div style="text-align:center;font-size:11px;color:${T.textMuted};font-weight:600">${w}</div>`).join('')}
          </div>
          <div style="display:grid;grid-template-columns:repeat(7,1fr);gap:8px">
            ${cells.map((d) => {
              if (!d) return '<div></div>';
              const key = `${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
              const { todoCount } = this.itemsForDate(key);
              const dayEvents = s.events.filter((e) => this.eventOccursOnDate(e, key));
              const previewLimit = 2;
              const shownEvents = dayEvents.slice(0, previewLimit);
              const overflowCount = dayEvents.length - shownEvents.length;
              const isToday = key === todayKeyStr;
              const border = isToday ? '1.5px solid ${T.accentSolid}' : `1px solid ${T.cardBorder}`;
              return `<div onclick="App.goToDayFromCalendar('${key}')" style="min-height:76px;border-radius:9px;padding:5px;display:flex;flex-direction:column;cursor:pointer;box-sizing:border-box;border:${border};background:${T.subtleBg || 'transparent'};overflow:hidden">
                <div style="display:flex;align-items:center;justify-content:space-between">
                  <div style="font-size:12px;font-weight:${isToday ? '700' : '500'};color:${isToday ? T.accentText : T.text}">${d}</div>
                  ${todoCount > 0 ? `<div title="To-dos" style="width:5px;height:5px;border-radius:50%;background:${T.accentText};flex:none"></div>` : ''}
                </div>
                <div style="display:flex;flex-direction:column;gap:2px;margin-top:4px">
                  ${shownEvents.map((e) => `<div style="font-size:9px;font-weight:600;line-height:1.3;color:${e.important ? T.orangeText : T.accentText};background:${e.important ? T.orangeSoftBg : T.accentSoftBg};border-radius:4px;padding:1.5px 4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${esc(e.name)}</div>`).join('')}
                  ${overflowCount > 0 ? `<div style="font-size:8.5px;color:${T.textMuted};padding-left:3px">+${overflowCount} more</div>` : ''}
                </div>
              </div>`;
            }).join('')}
          </div>
        </div>
        <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;overflow:hidden">
          <div style="padding:18px 22px;border-bottom:1px solid ${T.divider};font-weight:700;font-size:15px">Events this month</div>
          ${eventsThisMonth.length === 0 ? `<div style="padding:22px;font-size:13.5px;color:${T.textMuted}">No events this month.</div>` : eventsThisMonth.map((e) => `
            <div onclick="App.goToDayFromCalendar('${e.date}')" style="display:flex;align-items:center;gap:14px;padding:14px 22px;border-bottom:1px solid ${T.rowDivider};cursor:pointer">
              ${e.important ? `<div title="Important" style="width:8px;height:8px;border-radius:50%;background:${T.orangeText};flex:none"></div>` : ''}
              <div style="flex:1">
                <div style="font-size:14px;font-weight:500;color:${T.text}">${esc(e.name)}</div>
                ${e.source === 'ai' ? `<div style="font-size:11.5px;color:${T.textMuted};margin-top:2px">Added from mail</div>` : ''}
              </div>
              <div style="font-size:12px;color:${T.textMuted}">${formatTaskDate(e.date)}${e.endDate ? ` – ${formatTaskDate(e.endDate)}` : ''}</div>
              <div class="remove-btn" onclick="event.stopPropagation();App.removeEvent(${e.id})" style="width:22px;height:22px;border-radius:6px;display:flex;align-items:center;justify-content:center;cursor:pointer;color:${T.textMuted};flex:none">
                <svg width="12" height="12" viewBox="0 0 24 24" fill="none"><path d="M6 6l12 12M18 6L6 18" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"></path></svg>
              </div>
            </div>`).join('')}
        </div>`;
  },

  renderCalendarYearView(T) {
    const s = this.state;
    const year = s.calendarViewYear;
    const todayKeyStr = dayKey(new Date());
    const monthsInYear = [];
    for (let month = 0; month < 12; month++) {
      const firstOfMonth = new Date(year, month, 1);
      const startWeekday = firstOfMonth.getDay();
      const daysInMonth = new Date(year, month + 1, 0).getDate();
      const cells = [];
      for (let i = 0; i < startWeekday; i++) cells.push(null);
      for (let d = 1; d <= daysInMonth; d++) cells.push(d);
      monthsInYear.push({ month, label: firstOfMonth.toLocaleDateString(undefined, { month: 'long' }), cells });
    }
    return `
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px">
          <div onclick="App.calendarShiftYear(-1)" style="cursor:pointer;padding:4px 12px;font-size:15px;color:${T.textMuted}">←</div>
          <div style="font-weight:700;font-size:16px">${year}</div>
          <div onclick="App.calendarShiftYear(1)" style="cursor:pointer;padding:4px 12px;font-size:15px;color:${T.textMuted}">→</div>
        </div>
        <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px">
          ${monthsInYear.map((mo) => `
            <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:12px;padding:14px">
              <div onclick="App.calendarGoToMonth(${year},${mo.month})" style="font-weight:700;font-size:13px;margin-bottom:8px;cursor:pointer;color:${T.accentText}">${mo.label}</div>
              <div style="display:grid;grid-template-columns:repeat(7,1fr);gap:2px">
                ${mo.cells.map((d) => {
                  if (!d) return '<div></div>';
                  const key = `${year}-${String(mo.month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
                  const { total } = this.itemsForDate(key);
                  const isToday = key === todayKeyStr;
                  return `<div onclick="App.goToDayFromCalendar('${key}')" style="aspect-ratio:1;display:flex;align-items:center;justify-content:center;cursor:pointer;font-size:9.5px;font-weight:${isToday ? '700' : '400'};color:${isToday ? T.accentText : T.textMuted};border-radius:3px;background:${isToday ? T.accentSoftBg : 'transparent'};position:relative">${d}${total > 0 ? `<div style="position:absolute;bottom:1px;width:3px;height:3px;border-radius:50%;background:${T.accentText}"></div>` : ''}</div>`;
                }).join('')}
              </div>
            </div>`).join('')}
        </div>`;
  },


  renderSettings(T) {
    const s = this.state;
    const toggles = [
      { key: 'autoMode', label: 'Auto mode', description: 'Skip breaks and cycles — the timer counts up freely while you focus, and logs the exact time spent once you end the session.' },
    ];
    const themeOpts = ['light', 'dark', 'system'].map((th) => `
      <div onclick="App.selectTheme('${th}')" style="flex:1;text-align:center;padding:10px 0;border-radius:9px;border:1.5px solid ${s.theme === th ? T.accentSolid : T.unselectedBorder};background:${s.theme === th ? T.accentSolid : T.cardBg};color:${s.theme === th ? 'white' : T.text};cursor:pointer;font-weight:600;font-size:13.5px">${th[0].toUpperCase()}${th.slice(1)}</div>`).join('');
    return `
      <div id="scrollArea" style="flex:1;overflow-y:auto;padding:40px 48px;box-sizing:border-box;${this._animate ? 'animation:fadeIn .3s ease' : ''}">
        <div style="margin-bottom:32px">
          <div style="font-size:13px;color:${T.textMuted};font-weight:500;margin-bottom:4px">Preferences</div>
          <div style="font-family:'Manrope',sans-serif;font-weight:800;font-size:28px;letter-spacing:-0.02em">Settings</div>
        </div>
        <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;overflow:hidden;margin-bottom:20px">
          <div style="padding:18px 22px;border-bottom:1px solid ${T.divider};font-weight:700;font-size:15px">Account</div>
          <div style="display:flex;align-items:center;gap:14px;padding:18px 22px">
            <div style="width:44px;height:44px;border-radius:50%;background:oklch(0.7 0.12 152);display:flex;align-items:center;justify-content:center;color:white;font-weight:700;font-size:16px;flex:none">${esc(s.userInitial)}</div>
            <div style="flex:1">
              <div style="font-size:14.5px;font-weight:600">${esc(s.userName)}</div>
              <div style="font-size:12.5px;color:${T.textMuted}">Connected via Microsoft account</div>
            </div>
            <div onclick="App.signOut()" style="font-size:12.5px;font-weight:600;color:oklch(0.55 0.15 30);cursor:pointer;padding:8px 14px;border-radius:8px;border:1.5px solid oklch(0.55 0.1 30 / 0.35)">Sign out</div>
          </div>
        </div>
        <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;overflow:hidden;margin-bottom:20px">
          <div style="padding:18px 22px;border-bottom:1px solid ${T.divider};font-weight:700;font-size:15px">Focus sessions</div>
          ${toggles.map((tg) => {
            const on = s.settingsToggles[tg.key];
            return `
            <div style="display:flex;align-items:center;justify-content:space-between;padding:16px 22px;border-bottom:1px solid ${T.rowDivider}">
              <div>
                <div style="font-size:14px;font-weight:600;margin-bottom:2px">${tg.label}</div>
                <div style="font-size:12.5px;color:${T.textMuted}">${tg.description}</div>
              </div>
              <div onclick="App.toggleSetting('${tg.key}')" style="width:42px;height:24px;border-radius:12px;background:${on ? '${T.accentSolid}' : T.unselectedBorder};flex:none;cursor:pointer;padding:3px;box-sizing:border-box;display:flex;justify-content:${on ? 'flex-end' : 'flex-start'}">
                <div style="width:18px;height:18px;border-radius:50%;background:white;box-shadow:0 1px 2px rgba(0,0,0,0.15)"></div>
              </div>
            </div>`;
          }).join('')}
        </div>
        <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;overflow:hidden;margin-bottom:20px">
          <div style="padding:18px 22px;border-bottom:1px solid ${T.divider};font-weight:700;font-size:15px">Task categories</div>
          ${s.categories.map((cat, i) => `
            <div style="display:flex;align-items:center;gap:10px;padding:12px 22px;border-bottom:1px solid ${T.rowDivider}">
              <input value="${esc(cat)}" onchange="App.renameCategory(${i}, this.value)" style="flex:1;border:1px solid ${T.inputBorder};border-radius:8px;padding:8px 10px;font-size:13.5px;font-family:inherit;color:${T.text};background:transparent">
              <div class="remove-btn" onclick="App.removeCategory(${i})" style="width:26px;height:26px;border-radius:7px;display:flex;align-items:center;justify-content:center;cursor:pointer;color:${T.textMuted};flex:none">
                <svg width="12" height="12" viewBox="0 0 24 24" fill="none"><path d="M6 6l12 12M18 6L6 18" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"></path></svg>
              </div>
            </div>`).join('')}
          <div style="display:flex;align-items:center;gap:8px;padding:14px 22px">
            <input id="newCategoryInput" onkeydown="App.handleCategoryKeyDown(event)" placeholder="Add a category…" style="flex:1;border:1.5px dashed ${T.inputBorder};border-radius:10px;padding:10px 12px;font-size:14px;font-family:inherit;outline:none;color:${T.text};background:transparent">
            <div onclick="App.addCategory()" style="font-size:12.5px;font-weight:600;color:${T.accentText};cursor:pointer;padding:10px 14px;border-radius:8px;background:${T.accentSoftBg}">Add</div>
          </div>
        </div>
        <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:14px;overflow:hidden">
          <div style="padding:18px 22px;border-bottom:1px solid ${T.divider};font-weight:700;font-size:15px">Appearance</div>
          <div style="padding:18px 22px">
            <div style="font-size:12.5px;font-weight:600;color:${T.textMuted};margin-bottom:10px;text-transform:uppercase;letter-spacing:0.04em">Theme</div>
            <div style="display:flex;gap:8px">${themeOpts}</div>
          </div>
        </div>
      </div>`;
  },

  renderSetup(T) {
    const s = this.state;
    const openTasks = s.tasks.filter((t) => !t.done);
    return `
      <div id="scrollArea" style="flex:1;overflow-y:auto;display:flex;align-items:flex-start;justify-content:center;padding:40px 20px;box-sizing:border-box;${this._animate ? 'animation:fadeIn .3s ease' : ''}">
        <div style="width:460px;background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:18px;padding:36px;box-sizing:border-box">
          <div onclick="App.backToHome()" style="font-size:13px;font-weight:600;color:${T.textMuted};cursor:pointer;margin-bottom:20px;display:flex;align-items:center;gap:6px">← Back</div>
          <div style="font-family:'Manrope',sans-serif;font-weight:800;font-size:22px;letter-spacing:-0.01em;margin-bottom:6px">Set up your session</div>
          <div style="font-size:13.5px;color:${T.textMuted};margin-bottom:28px">Pick a task, how long to focus, and whether to take breaks.</div>
          <div style="font-size:12.5px;font-weight:600;color:${T.textMuted};margin-bottom:10px;text-transform:uppercase;letter-spacing:0.04em">Task</div>
          <div style="display:flex;flex-direction:column;gap:8px;margin-bottom:24px">
            ${openTasks.map((t) => `
              <div style="padding:4px 4px 4px 14px;border-radius:10px;border:1.5px solid ${t.id === s.selectedTaskId ? '${T.accentSolid}' : T.unselectedBorder};background:${t.id === s.selectedTaskId ? (T.dark ? 'oklch(0.28 0.05 152)' : 'oklch(0.96 0.03 152)') : T.cardBg}">
                <div style="display:flex;align-items:center;gap:8px">
                  <div onclick="App.selectTask(${t.id})" style="flex:1;padding:8px 0;cursor:pointer;font-size:14px;font-weight:500;display:flex;justify-content:space-between;align-items:center">
                    <span>${esc(t.name)}${t.date ? ` <span style="font-size:11.5px;font-weight:600;color:${T.textMuted}">· ${formatTaskDate(t.date)}</span>` : ''}</span>
                    ${t.id === s.selectedTaskId ? '<div style="width:16px;height:16px;border-radius:50%;background:${T.accentSolid};display:flex;align-items:center;justify-content:center;margin-right:8px"><svg width="9" height="9" viewBox="0 0 24 24" fill="none"><path d="M5 13l4 4L19 7" stroke="white" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"></path></svg></div>' : ''}
                  </div>
                  <div class="remove-btn" onclick="App.removeSetupTask(${t.id})" style="width:26px;height:26px;border-radius:7px;display:flex;align-items:center;justify-content:center;cursor:pointer;color:${T.textMuted};flex:none">
                    <svg width="12" height="12" viewBox="0 0 24 24" fill="none"><path d="M6 6l12 12M18 6L6 18" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"></path></svg>
                  </div>
                </div>
                ${(t.subtasks || []).length > 0 ? `
                  <div style="padding:2px 8px 8px 4px;display:flex;flex-direction:column;gap:4px">
                    ${t.subtasks.map((st) => `
                      <div style="display:flex;align-items:center;gap:8px">
                        <div onclick="App.toggleSubtaskDone(${t.id},${st.id})" style="width:13px;height:13px;border-radius:4px;border:1.5px solid ${st.done ? '${T.accentSolid}' : T.unselectedBorder};background:${st.done ? '${T.accentSolid}' : 'transparent'};flex:none;cursor:pointer"></div>
                        <div style="font-size:12.5px;font-weight:500;text-decoration:${st.done ? 'line-through' : 'none'};color:${T.textMuted}">${esc(st.name)}${st.date ? ` · ${formatTaskDate(st.date)}` : ''}</div>
                      </div>`).join('')}
                  </div>` : ''}
              </div>`).join('')}
          </div>
          <div style="margin-bottom:28px">
            <div style="display:flex;align-items:center;gap:8px">
              <input id="newSetupTaskInput" onkeydown="App.handleSetupTaskKeyDown(event)" placeholder="Add your own task to focus on…" style="flex:1;border:1.5px dashed ${T.inputBorder};border-radius:10px;padding:10px 12px;font-size:14px;font-family:inherit;outline:none;color:${T.text};background:transparent">
              <select id="newSetupTaskCategory" style="border:1px solid ${T.cardBorder};border-radius:8px;padding:9px 8px;font-size:12.5px;font-family:inherit;color:${T.text};background:${T.cardBg};cursor:pointer;color-scheme:${T.dark ? 'dark' : 'light'}">
                ${s.categories.map((c) => `<option value="${c}">${c}</option>`).join('')}
              </select>
              <div onclick="App.addSetupTask()" style="font-size:12.5px;font-weight:600;color:${T.accentText};cursor:pointer;padding:10px 14px;border-radius:8px;background:${T.accentSoftBg}">Add</div>
            </div>
            <div style="padding-left:4px;margin-top:8px;display:flex;flex-direction:column;gap:6px">
              ${s.newSetupTaskDraftSubtasks.map((name, i) => `
                <div style="display:flex;align-items:center;gap:8px">
                  <div style="font-size:12.5px;color:${T.textMuted};flex:1">— ${esc(name)}</div>
                  <div class="remove-btn" onclick="App.removeSetupDraftSubtask(${i})" style="width:18px;height:18px;border-radius:5px;display:flex;align-items:center;justify-content:center;cursor:pointer;color:${T.textMuted};flex:none">
                    <svg width="10" height="10" viewBox="0 0 24 24" fill="none"><path d="M6 6l12 12M18 6L6 18" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"></path></svg>
                  </div>
                </div>`).join('')}
              <div style="display:flex;align-items:center;gap:8px">
                <input id="newSetupTaskDraftSubtaskInput" onkeydown="App.handleSetupDraftSubtaskKeyDown(event)" placeholder="+ Add a step for this task (optional)" style="flex:1;border:1.5px dashed ${T.inputBorder};border-radius:8px;padding:7px 10px;font-size:12.5px;font-family:inherit;outline:none;color:${T.text};background:transparent">
                <div onclick="App.addSetupDraftSubtask()" style="font-size:12px;font-weight:600;color:${T.accentText};cursor:pointer;padding:6px 12px;border-radius:7px;background:${T.accentSoftBg}">Add step</div>
              </div>
            </div>
          </div>

          ${s.settingsToggles.autoMode ? `
          <div style="padding:16px;border-radius:10px;background:${T.subtleBg};margin-bottom:28px;font-size:13px;color:${T.textMuted};line-height:1.5">
            Auto mode is on (Settings → Focus sessions) — this session will count up with no set length, breaks, or cycles. It ends when you end it, and logs the real time you focused.
          </div>` : `
          <div style="display:flex;gap:24px;margin-bottom:28px">
            <div style="flex:1;text-align:center">
              <div style="font-size:12.5px;font-weight:600;color:${T.textMuted};margin-bottom:10px;text-transform:uppercase;letter-spacing:0.04em">Focus length</div>
              ${this.renderRoller({ id: 'focusRoller', target: 'focusMinutes', min: 1, max: 120, step: 1, value: s.focusMinutes, unit: 'm' })}
            </div>
            ${s.breakEnabled ? `
            <div style="flex:1;text-align:center">
              <div style="font-size:12.5px;font-weight:600;color:${T.textMuted};margin-bottom:10px;text-transform:uppercase;letter-spacing:0.04em">Break length</div>
              ${this.renderRoller({ id: 'breakRoller', target: 'breakMinutes', min: 1, max: 30, step: 1, value: s.breakMinutes, unit: 'm' })}
            </div>` : ''}
          </div>

          <div style="display:flex;align-items:center;justify-content:space-between;padding:14px 16px;border-radius:10px;background:${T.subtleBg};margin-bottom:20px">
            <div>
              <div style="font-size:14px;font-weight:600;margin-bottom:2px">Take breaks between sessions</div>
              <div style="font-size:12.5px;color:${T.textMuted}">A break follows every focus session except the last</div>
            </div>
            <div onclick="App.toggleBreakEnabled()" style="width:42px;height:24px;border-radius:12px;background:${s.breakEnabled ? '${T.accentSolid}' : T.unselectedBorder};flex:none;cursor:pointer;padding:3px;box-sizing:border-box;display:flex;justify-content:${s.breakEnabled ? 'flex-end' : 'flex-start'}">
              <div style="width:18px;height:18px;border-radius:50%;background:white;box-shadow:0 1px 2px rgba(0,0,0,0.15)"></div>
            </div>
          </div>

          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:32px">
            <div>
              <div style="font-size:12.5px;font-weight:600;color:${T.textMuted};text-transform:uppercase;letter-spacing:0.04em">Cycles</div>
              <div style="font-size:12.5px;color:${T.textMuted};margin-top:2px">1 cycle = 1 focus session${s.breakEnabled ? ' + break' : ''}</div>
            </div>
            <div style="display:flex;align-items:center;gap:14px">
              <div onclick="App.changeCycles(-1)" style="width:32px;height:32px;border-radius:9px;border:1.5px solid ${T.unselectedBorder};display:flex;align-items:center;justify-content:center;cursor:pointer;font-weight:700;font-size:16px">−</div>
              <div style="font-family:'Manrope',sans-serif;font-weight:800;font-size:18px;min-width:20px;text-align:center">${s.cycles}</div>
              <div onclick="App.changeCycles(1)" style="width:32px;height:32px;border-radius:9px;border:1.5px solid ${T.unselectedBorder};display:flex;align-items:center;justify-content:center;cursor:pointer;font-weight:700;font-size:16px">+</div>
            </div>
          </div>`}

          <div class="primary-btn" onclick="App.startSession()" style="width:100%;box-sizing:border-box;text-align:center;padding:14px 0;background:${T.accentSolid};color:white;border-radius:11px;font-weight:700;font-size:15px;cursor:pointer;box-shadow:0 4px 14px ${T.dark ? 'oklch(0.52 0.15 152 / 0.32)' : 'oklch(0.36 0.09 152 / 0.28)'}">Start focusing</div>
        </div>
      </div>`;
  },

  renderTimer() {
    const s = this.state;
    const task = s.tasks.find((t) => t.id === s.timerTaskId) || s.tasks[0] || { name: 'Focus session' };
    const isBreak = s.currentPhase === 'break';
    const isAuto = s.currentPhase === 'focus-auto';
    const circumference = 791.7;
    let dashOffset, display;
    if (isAuto) {
      // No target duration — loop the ring once a minute just to show motion,
      // and count the clock up instead of down.
      const loopProgress = (s.phaseElapsedSeconds % 60) / 60;
      dashOffset = circumference * (1 - loopProgress);
      const mm = Math.floor(s.phaseElapsedSeconds / 60);
      const ss = s.phaseElapsedSeconds % 60;
      display = `${String(mm).padStart(2, '0')}:${String(ss).padStart(2, '0')}`;
    } else {
      const totalSeconds = s.phaseDurationMin * 60 || 1;
      const progress = s.remainingSeconds / totalSeconds;
      dashOffset = circumference * (1 - progress);
      const mm = Math.floor(s.remainingSeconds / 60);
      const ss = s.remainingSeconds % 60;
      display = `${String(mm).padStart(2, '0')}:${String(ss).padStart(2, '0')}`;
    }
    const cycleLabel = (!isAuto && s.cycles > 1) ? `Cycle ${s.currentCycleIndex + 1} of ${s.cycles}` : '';
    return `
      <div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;background:${isBreak ? 'oklch(0.24 0.05 220)' : 'oklch(0.22 0.02 152)'};color:white;${this._animate ? 'animation:fadeIn .3s ease' : ''}">
        ${cycleLabel ? `<div style="font-size:12px;font-weight:600;color:oklch(0.8 0.03 152 / 0.75);margin-bottom:6px">${cycleLabel}</div>` : ''}
        ${isAuto ? `<div style="font-size:12px;font-weight:600;color:oklch(0.8 0.03 152 / 0.75);margin-bottom:6px">Auto mode — counting up</div>` : ''}
        <div style="font-size:13px;font-weight:600;letter-spacing:0.06em;text-transform:uppercase;color:${isBreak ? 'oklch(0.8 0.09 220)' : 'oklch(0.75 0.05 152)'};margin-bottom:16px">${isBreak ? 'Break' : esc(task.name)}</div>
        <div style="position:relative;width:280px;height:280px;display:flex;align-items:center;justify-content:center;margin-bottom:36px">
          <svg width="280" height="280" style="position:absolute;transform:rotate(-90deg)">
            <circle cx="140" cy="140" r="126" fill="none" stroke="oklch(0.32 0.02 152)" stroke-width="10"></circle>
            <circle cx="140" cy="140" r="126" fill="none" stroke="${isBreak ? 'oklch(0.72 0.12 220)' : 'oklch(0.65 0.15 152)'}" stroke-width="10" stroke-linecap="round" stroke-dasharray="791.7" stroke-dashoffset="${dashOffset}" style="transition:stroke-dashoffset 1s linear"></circle>
          </svg>
          <div style="font-family:'Manrope',sans-serif;font-size:52px;font-weight:800;letter-spacing:-0.02em">${display}</div>
        </div>
        <div style="display:flex;gap:14px">
          <div onclick="App.togglePause()" style="padding:13px 28px;border-radius:11px;background:white;color:oklch(0.22 0.02 152);font-weight:700;font-size:14.5px;cursor:pointer">${s.isPaused ? 'Resume' : 'Pause'}</div>
          ${isBreak ? `<div onclick="App.skipBreak()" style="padding:13px 28px;border-radius:11px;background:oklch(0.3 0.02 152);color:oklch(0.85 0.02 152);font-weight:700;font-size:14.5px;cursor:pointer;border:1px solid oklch(0.4 0.02 152)">Skip break</div>` : ''}
          <div onclick="App.endSession()" style="padding:13px 28px;border-radius:11px;background:oklch(0.3 0.02 152);color:oklch(0.85 0.02 152);font-weight:700;font-size:14.5px;cursor:pointer;border:1px solid oklch(0.4 0.02 152)">${isAuto ? "I'm done" : 'End session'}</div>
        </div>
      </div>`;
  },

  renderCheckin(T) {
    const s = this.state;
    const task = s.tasks.find((t) => t.id === s.timerTaskId) || { name: 'this task' };
    return `
      <div style="flex:1;display:flex;align-items:center;justify-content:center;${this._animate ? 'animation:fadeIn .3s ease' : ''}">
        <div style="width:400px;display:flex;flex-direction:column;align-items:center;text-align:center">
          <div style="font-family:'Manrope',sans-serif;font-weight:800;font-size:24px;letter-spacing:-0.01em;margin-bottom:10px">Focus session done</div>
          <div style="font-size:14.5px;color:${T.textMuted};line-height:1.5;margin-bottom:28px">Did you finish <strong style="color:${T.text}">${esc(task.name)}</strong>?</div>
          <div style="display:flex;gap:10px;width:100%">
            <div onclick="App.checkinTaskNotDone()" style="flex:1;text-align:center;padding:13px 0;border-radius:11px;border:1.5px solid ${T.cardBorder};font-weight:600;font-size:14px;cursor:pointer">Not yet</div>
            <div class="primary-btn" onclick="App.checkinTaskDone()" style="flex:1;text-align:center;padding:13px 0;border-radius:11px;background:${T.accentSolid};color:white;font-weight:600;font-size:14px;cursor:pointer">Yes, done!</div>
          </div>
        </div>
      </div>`;
  },

  renderBreakdownPrompt(T) {
    const s = this.state;
    const task = s.tasks.find((t) => t.id === s.timerTaskId) || { name: 'this task' };
    return `
      <div style="flex:1;display:flex;align-items:center;justify-content:center;${this._animate ? 'animation:fadeIn .3s ease' : ''}">
        <div style="width:400px;display:flex;flex-direction:column;align-items:center;text-align:center">
          <div style="font-family:'Manrope',sans-serif;font-weight:800;font-size:24px;letter-spacing:-0.01em;margin-bottom:10px">No worries.</div>
          <div style="font-size:14.5px;color:${T.textMuted};line-height:1.5;margin-bottom:28px">Want to break <strong style="color:${T.text}">${esc(task.name)}</strong> into smaller steps?</div>
          <div style="display:flex;gap:10px;width:100%">
            <div onclick="App.declineBreakdown()" style="flex:1;text-align:center;padding:13px 0;border-radius:11px;border:1.5px solid ${T.cardBorder};font-weight:600;font-size:14px;cursor:pointer">No, continue</div>
            <div class="primary-btn" onclick="App.acceptBreakdown()" style="flex:1;text-align:center;padding:13px 0;border-radius:11px;background:${T.accentSolid};color:white;font-weight:600;font-size:14px;cursor:pointer">Yes, break it down</div>
          </div>
        </div>
      </div>`;
  },

  renderBreakdownEdit(T) {
    const s = this.state;
    const task = s.tasks.find((t) => t.id === s.timerTaskId) || { name: 'this task', subtasks: [] };
    const subtasks = task.subtasks || [];
    return `
      <div style="flex:1;display:flex;align-items:center;justify-content:center;${this._animate ? 'animation:fadeIn .3s ease' : ''}">
        <div style="width:460px;background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:18px;padding:36px;box-sizing:border-box">
          <div style="font-family:'Manrope',sans-serif;font-weight:800;font-size:22px;letter-spacing:-0.01em;margin-bottom:6px">Break down the task</div>
          <div style="font-size:13.5px;color:${T.textMuted};margin-bottom:24px">Add smaller steps for <strong style="color:${T.text}">${esc(task.name)}</strong>.</div>
          <div style="display:flex;flex-direction:column;gap:8px;margin-bottom:16px">
            ${subtasks.length === 0 ? `<div style="font-size:13.5px;color:${T.textMuted}">No steps added yet.</div>` : subtasks.map((st) => `
              <div style="display:flex;align-items:center;gap:10px;padding:10px 12px;border-radius:9px;border:1px solid ${T.cardBorder}">
                <div onclick="App.toggleSubtaskDone(${task.id},${st.id})" style="width:17px;height:17px;border-radius:5px;border:1.5px solid ${st.done ? '${T.accentSolid}' : T.unselectedBorder};background:${st.done ? '${T.accentSolid}' : T.cardBg};flex:none;cursor:pointer"></div>
                <div style="flex:1;font-size:13.5px;font-weight:500;text-decoration:${st.done ? 'line-through' : 'none'};color:${st.done ? T.textMuted : T.text}">${esc(st.name)}</div>
                <div class="remove-btn" onclick="App.removeSubtask(${task.id},${st.id})" style="width:22px;height:22px;border-radius:6px;display:flex;align-items:center;justify-content:center;cursor:pointer;color:${T.textMuted};flex:none">
                  <svg width="11" height="11" viewBox="0 0 24 24" fill="none"><path d="M6 6l12 12M18 6L6 18" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"></path></svg>
                </div>
              </div>`).join('')}
          </div>
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:26px">
            <input id="newSubtaskInput" onkeydown="App.handleSubtaskKeyDown(event,${task.id},'newSubtaskInput')" placeholder="Add a smaller step…" style="flex:1;border:1.5px dashed ${T.inputBorder};border-radius:10px;padding:10px 12px;font-size:14px;font-family:inherit;outline:none;color:${T.text};background:transparent">
            <div onclick="App.addSubtaskToTask(${task.id},'newSubtaskInput')" style="font-size:12.5px;font-weight:600;color:${T.accentText};cursor:pointer;padding:10px 14px;border-radius:8px;background:${T.accentSoftBg}">Add</div>
          </div>
          <div class="primary-btn" onclick="App.finishBreakdown()" style="width:100%;box-sizing:border-box;text-align:center;padding:14px 0;background:${T.accentSolid};color:white;border-radius:11px;font-weight:700;font-size:15px;cursor:pointer">Continue</div>
        </div>
      </div>`;
  },

  renderComplete(T) {
    const s = this.state;
    const task = s.tasks.find((t) => t.id === s.timerTaskId) || s.tasks[0] || { name: 'your task' };
    const cycleNote = s.lastSessionCycles > 1 ? ` across ${s.lastSessionCycles} cycles` : '';
    return `
      <div style="flex:1;display:flex;align-items:center;justify-content:center;${this._animate ? 'animation:fadeIn .3s ease' : ''}">
        <div style="width:400px;display:flex;flex-direction:column;align-items:center;text-align:center">
          <div style="width:64px;height:64px;border-radius:50%;background:${T.accentSoftBg};display:flex;align-items:center;justify-content:center;margin-bottom:22px;animation:popIn .5s cubic-bezier(.2,.8,.2,1)">
            <svg width="28" height="28" viewBox="0 0 24 24" fill="none"><path d="M5 13l4 4L19 7" stroke="oklch(0.5 0.15 152)" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"></path></svg>
          </div>
          <div style="font-family:'Manrope',sans-serif;font-weight:800;font-size:24px;letter-spacing:-0.01em;margin-bottom:8px">Nice focus.</div>
          <div style="font-size:14.5px;color:${T.textMuted};line-height:1.5;margin-bottom:28px">You spent <strong style="color:${T.text}">${formatDuration(s.lastSessionFocusSeconds)}</strong> focused on <strong style="color:${T.text}">${esc(task.name)}</strong>${cycleNote}.</div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;width:100%;margin-bottom:28px">
            <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:12px;padding:16px">
              <div style="font-size:12px;color:${T.textMuted};margin-bottom:6px">Sessions today</div>
              <div style="font-family:'Manrope',sans-serif;font-weight:800;font-size:20px">${s.sessionsToday}</div>
            </div>
            <div style="background:${T.cardBg};border:1px solid ${T.cardBorder};border-radius:12px;padding:16px">
              <div style="font-size:12px;color:${T.textMuted};margin-bottom:6px">Current streak</div>
              <div style="font-family:'Manrope',sans-serif;font-weight:800;font-size:20px">${s.streak} days</div>
            </div>
          </div>
          <div style="display:flex;gap:10px;width:100%">
            <div onclick="App.backToHome()" style="flex:1;text-align:center;padding:13px 0;border-radius:11px;border:1.5px solid ${T.cardBorder};font-weight:600;font-size:14px;cursor:pointer">Back to home</div>
            <div class="primary-btn" onclick="App.goToSetup()" style="flex:1;text-align:center;padding:13px 0;border-radius:11px;background:${T.accentSolid};color:white;font-weight:600;font-size:14px;cursor:pointer">Focus again</div>
          </div>
        </div>
      </div>`;
  },

  render() {
    const s = this.state;
    // Home/Stats always read live off the real persisted history, never off
    // hand-maintained counters, so they can't drift from what actually happened.
    const todayRec = this.todayRecord();
    s.totalFocusSecondsToday = todayRec.focusSeconds;
    s.sessionsToday = todayRec.sessionsCompleted;
    s.streak = this.currentStreak();
    const T = this.theme();
    // Only replay the entrance animation when the screen actually changes —
    // otherwise every click/tick fully rebuilds the DOM and restarts
    // fadeIn's transform, which looked like the whole page flashing.
    this._animate = this._lastScreen !== s.screen;
    // Rebuilding #root's innerHTML always creates a fresh scroll container at
    // scrollTop 0 — so any edit while scrolled down snapped back to the top.
    // Carry the old scroll position forward when we're staying on the same
    // screen (a real screen change should still start at the top).
    const prevScrollArea = document.getElementById('scrollArea');
    const savedScrollTop = (!this._animate && prevScrollArea) ? prevScrollArea.scrollTop : 0;
    this._lastScreen = s.screen;
    const showSidebar = ['home', 'setup', 'complete', 'stats', 'inbox', 'calendar', 'settings'].includes(s.screen);
    let body = '';
    if (s.screen === 'onboarding') body = this.renderOnboarding(T);
    else if (s.screen === 'home') body = this.renderHome(T);
    else if (s.screen === 'setup') body = this.renderSetup(T);
    else if (s.screen === 'timer') body = this.renderTimer(T);
    else if (s.screen === 'checkin') body = this.renderCheckin(T);
    else if (s.screen === 'breakdown-prompt') body = this.renderBreakdownPrompt(T);
    else if (s.screen === 'breakdown-edit') body = this.renderBreakdownEdit(T);
    else if (s.screen === 'complete') body = this.renderComplete(T);
    else if (s.screen === 'stats') body = this.renderStats(T);
    else if (s.screen === 'inbox') body = this.renderInbox(T);
    else if (s.screen === 'calendar') body = this.renderCalendar(T);
    else if (s.screen === 'settings') body = this.renderSettings(T);

    document.getElementById('root').innerHTML = `
      <div style="width:100%;height:100vh;background:${T.pageBg};font-family:'Inter',system-ui,sans-serif;color:${T.text};display:flex;overflow:hidden;color-scheme:${T.dark ? 'dark' : 'light'};--hover-bg:${T.hoverBg};--row-hover-bg:${T.hoverBg}">
        ${showSidebar ? this.renderSidebar(T) : ''}
        <div style="flex:1;min-width:0;position:relative;display:flex;flex-direction:column">${body}</div>
      </div>`;
    const newScrollArea = document.getElementById('scrollArea');
    if (newScrollArea && savedScrollTop) newScrollArea.scrollTop = savedScrollTop;
    this.attachRollerListeners();

    // Cheap, so we just do it on every render rather than chasing every
    // individual mutation site — keeps tasks/events/reminders from
    // vanishing on refresh the way they used to.
    this.saveTasks();
    this.saveEvents();
    this.saveReminders();
    if (this.authProvider) this.maybeSyncToServer();
  },

  // ---------- Time roller (Setup screen) ----------

  renderRoller({ id, target, min, max, step = 1, value, unit = '' }) {
    const itemHeight = 30;
    const visibleCount = 5;
    const values = [];
    for (let v = min; v <= max; v += step) values.push(v);
    const padding = itemHeight * Math.floor(visibleCount / 2);
    return `
      <div style="position:relative;height:${itemHeight * visibleCount}px">
        <div id="${id}" class="roller" data-target="${target}" data-min="${min}" data-step="${step}" data-item-height="${itemHeight}" style="height:100%;overflow-y:scroll;scroll-snap-type:y mandatory">
          <div style="height:${padding}px"></div>
          ${values.map((v) => `<div class="roller-item" style="height:${itemHeight}px;scroll-snap-align:center;display:flex;align-items:center;justify-content:center;font-size:15px;font-weight:${v === value ? 700 : 500}">${v}${unit}</div>`).join('')}
          <div style="height:${padding}px"></div>
        </div>
        <div style="position:absolute;top:50%;left:0;right:0;height:${itemHeight}px;transform:translateY(-50%);border-top:1.5px solid currentColor;border-bottom:1.5px solid currentColor;opacity:0.15;pointer-events:none"></div>
      </div>`;
  },

  attachRollerListeners() {
    document.querySelectorAll('.roller').forEach((el) => {
      const itemHeight = parseFloat(el.dataset.itemHeight);
      const min = parseFloat(el.dataset.min);
      const step = parseFloat(el.dataset.step);
      const target = el.dataset.target;
      const index = Math.round((this.state[target] - min) / step);
      el.scrollTop = index * itemHeight;

      let debounceTimer;
      const settle = () => {
        const idx = Math.max(0, Math.round(el.scrollTop / itemHeight));
        this.state[target] = min + idx * step;
      };
      el.addEventListener('scroll', () => {
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(settle, 120);
      });
      el.addEventListener('scrollend', settle);
    });
  },
};

window.App = App;
document.addEventListener('DOMContentLoaded', () => {
  App.loadStats();
  App.loadCategories();
  App.loadTaskHistory();
  App.loadMoodLog();
  App.loadEmailNotes();
  App.loadHabits();
  App.loadHabitLog();
  App.loadTasks();
  App.loadEvents();
  App.loadReminders();
  App.render();
  App.initMsal().catch((err) => console.error('MSAL init failed:', err));
});

window.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'hidden' && App.authProvider) App.syncToServerNow();
});

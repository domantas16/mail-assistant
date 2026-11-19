import { CONFIG } from './config.js';
import {
  $, $$, esc, store, toast, debounce, avatar, fmtDate, fmtFull, fmtSize, textToHtml, htmlToText,
  cleanEmailHtml, srcdocFor, sanitizeBasic, copyRich, downloadBlob, grow, dayLabel, ICON,
} from './lib.js';
import * as auth from './auth.js';
import { graph } from './graph.js';
import { createDemo, cannedDraft, DEMO_SHARED } from './demo.js';
import * as ai from './ai.js';
import { DEFAULT_ANSWERS, GENERIC, parseAnswers, pickAnswer, fillAnswer, firstName } from './answers.js';

// ---------- settings & state ----------
const DEFAULTS = {
  clientId: '', tenantId: '', apiKey: '', model: '', effort: 'low', myName: '', myRole: '',
  style: '', notes: '', useExamples: true, signatureHtml: '', addSignature: true,
  draftFont: 'Aptos, Calibri, Arial, sans-serif', loadImages: false, sharedMailboxes: '',
  aiMode: 'full', autoDraft: true, answers: '', accSend: false, accShared: false, accContacts: false,
};
const cfg = () => ({ ...DEFAULTS, ...store.get('settings', {}) });
const setCfg = (patch) => store.set('settings', { ...store.get('settings', {}), ...patch });
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const S = {};
function resetState() {
  Object.assign(S, {
    mode: null, api: null, me: null, view: 'mail',
    folders: { main: [], other: [] }, folder: null,
    list: [], next: null, search: '', listBusy: false, listSeq: 0,
    sel: null, thread: [], drafts: new Map(), compose: null,
    contactsTab: 'frequent', contacts: null, people: [], contactQ: '', contact: null, contactSeq: 0,
    examples: null,
    chats: null, chat: null, chatQ: '', chatSeq: 0, mailbox: null, ready: new Map(), prepSeq: 0,
  });
}
resetState();
const E = { app: $('#app'), side: $('#sidebar'), list: $('#listPane'), detail: $('#detailPane') };

// ---------- what the app may do ----------
// Live: decided by the permissions Microsoft granted at sign-in. Demo: everything, to show the full app.
function can(f) {
  if (S.mode === 'demo') return true;
  const g = auth.grantedScopes();
  return {
    send: g.has('mail.send') && g.has('mail.readwrite'),
    shared: g.has('mail.read.shared') || g.has('mail.readwrite.shared'),
    sharedSend: g.has('mail.send.shared') && g.has('mail.readwrite.shared'),
    people: g.has('people.read'),
    contacts: g.has('contacts.read'),
  }[f] || false;
}
const canSendHere = () => can('send') && (!S.mailbox || can('sharedSend'));
const aiMode = () => cfg().aiMode || 'full';
const answers = () => parseAnswers(cfg().answers || DEFAULT_ANSWERS);
const BULK = /(^|[._-])(no-?reply|newsletter|news|events?|marketing|notifications?|mailer|info)@|@(mailchimp|hubspot|linkedin|.*\.mailer)/i;
function isCustomer(p) {
  const e = p?.email || '';
  if (!e || e === myAddr() || BULK.test(e)) return false;
  const dom = (x) => (x || '').split('@')[1] || '';
  return dom(e) !== dom(S.me?.email) && dom(e) !== dom(S.mailbox);
}

// ---------- mailboxes: your own plus shared ones (info@, sales@ ...) ----------
const EMAIL = /[^\s<>,;"]+@[^\s<>,;"]+\.[^\s<>,;"]+/g;
function mailboxes() {
  if (S.mode === 'demo') return [DEMO_SHARED];
  if (!can('shared')) return [];
  return [...new Set((cfg().sharedMailboxes.match(EMAIL) || []).map((x) => x.toLowerCase()))];
}
// The address that counts as "me" in the open mailbox (a shared mailbox sends as itself).
const myAddr = () => S.mailbox || S.me?.email || '';

function restoreMailbox() {
  const saved = store.get('mailbox');
  S.mailbox = saved && mailboxes().includes(saved) ? saved : null;
  S.api.setMailbox(S.mailbox);
}

async function switchMailbox(address) {
  for (const st of [...S.drafts.values(), S.compose].filter(Boolean)) {
    if (st.pending) { clearTimeout(st.pending.timer); toast('Sending cancelled because you switched mailbox.', 'err'); }
  }
  const keep = { mode: S.mode, api: S.api, me: S.me, examples: S.examples };
  resetState();
  Object.assign(S, keep);
  S.mailbox = address || null;
  store.set('mailbox', S.mailbox);
  S.api.setMailbox(S.mailbox);
  loadReady();
  E.app.classList.remove('show-detail');
  E.list.innerHTML = '';
  E.detail.innerHTML = loading(`Opening ${S.mailbox || 'your mailbox'}…`);
  await loadFolders();
  S.folder = S.folders.main[0] || { id: 'inbox', name: 'Inbox', key: 'inbox' };
  openChats();
}

// ---------- pinned contacts (per mailbox) ----------
const pinKey = () => `pins:${S.mailbox || 'me'}`;
const getPins = () => store.get(pinKey(), []);
const isPinned = (email) => getPins().some((p) => p.email === email);
function togglePin(person) {
  const pins = getPins();
  const next = isPinned(person.email) ? pins.filter((p) => p.email !== person.email) : [...pins, { name: person.name, email: person.email }];
  store.set(pinKey(), next);
  toast(isPinned(person.email) ? `Pinned ${person.name}` : `Unpinned ${person.name}`, 'ok', 1500);
}

// ---------- start ----------
async function boot() {
  wireEvents();
  let fresh = false;
  try {
    fresh = !!(await auth.handleRedirect());
  } catch (e) {
    return renderWelcome(e);
  }
  if (store.get('demo') && !fresh) return startDemo();
  if (auth.getTokens()) return startLive(fresh);
  renderWelcome();
}

async function startLive(fresh) {
  resetState();
  S.mode = 'live';
  S.api = graph;
  store.del('demo');
  E.app.className = 'app';
  E.side.innerHTML = '';
  E.list.innerHTML = '';
  E.detail.innerHTML = loading('Connecting to Outlook…');
  try {
    S.me = await graph.me();
  } catch (e) {
    return fail(e, true);
  }
  if (fresh) toast(`Signed in as ${S.me.email}`, 'ok');
  restoreMailbox();
  loadReady();
  await loadFolders();
  S.folder = S.folders.main[0] || { id: 'inbox', name: 'Inbox', key: 'inbox' };
  openChats();
}

async function startDemo() {
  resetState();
  S.mode = 'demo';
  S.api = createDemo();
  store.set('demo', true);
  E.app.className = 'app';
  S.me = await S.api.me();
  restoreMailbox();
  loadReady();
  await loadFolders();
  S.folder = S.folders.main[0];
  openChats();
}

function fail(e, fatal = false) {
  console.error(e);
  if (e?.code === 'auth_required') {
    auth.signOut();
    return renderWelcome(e);
  }
  toast(e?.message || String(e), 'err', 7000);
  if (fatal) renderWelcome(e);
}

// ---------- welcome / setup ----------
function friendlyAuthError(e) {
  const m = `${e?.code || ''} ${e?.message || e}`;
  if (/AADSTS65001|consent_required|admin approval|AADSTS90094/i.test(m)) return 'Microsoft says this app still needs admin approval. Ask IT to click "Grant admin consent" for Mail Assistant.';
  if (/AADSTS50011/i.test(m)) return `The redirect address doesn't match. IT must register exactly ${auth.redirectUri()} under "Single-page application".`;
  if (/AADSTS9002326|Single-Page Application/i.test(m)) return 'The redirect address is registered as "Web". IT needs to move it to "Single-page application".';
  if (/AADSTS700016|AADSTS90002|AADSTS900023/i.test(m)) return "Microsoft can't find this app. Check the Application (client) ID and Directory (tenant) ID.";
  if (/access_denied/i.test(m)) return 'Sign-in was cancelled or blocked.';
  return e?.message || String(e);
}

function renderWelcome(err) {
  resetState();
  E.app.className = 'app welcome';
  E.side.innerHTML = '';
  E.list.innerHTML = '';
  const c = cfg();
  const clientId = c.clientId || CONFIG.clientId;
  const tenantId = c.tenantId || CONFIG.tenantId;
  E.detail.innerHTML = `
  <div class="welcome-card">
    <div class="brand big">${ICON.logo}<span>Mail Assistant</span></div>
    <p class="lead">Your Outlook mail as chats, with a reply already drafted for each customer email. Check it and press Send.</p>
    ${err ? `<div class="alert err">${esc(friendlyAuthError(err))}</div>` : ''}
    <ol class="steps">
      <li class="step">
        <div class="step-h"><b>Connect Microsoft 365</b><span class="chip ${clientId ? 'ok' : 'wait'}">${clientId ? 'App IDs added' : 'Waiting for IT'}</span></div>
        <p class="muted">IT registers the app once and sends you two IDs. It starts read-only (User.Read, Mail.Read); sending and shared mailboxes can be switched on later. Paste the IDs here:</p>
        <div class="grid2">
          <label>Application (client) ID<input id="wClient" value="${esc(clientId)}" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" spellcheck="false" autocomplete="off"></label>
          <label>Directory (tenant) ID<input id="wTenant" value="${esc(tenantId)}" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" spellcheck="false" autocomplete="off"></label>
        </div>
        <p class="muted small">Redirect URI for IT: <code>${esc(auth.redirectUri())}</code> <button class="link" data-act="copy-text" data-text="${esc(auth.redirectUri())}">Copy</button></p>
        <button class="btn primary" data-act="signin">${ICON.mail}<span>Sign in with Microsoft</span></button>
      </li>
      <li class="step">
        <div class="step-h"><b>Add your Claude API key (optional)</b><span class="chip ${c.apiKey ? 'ok' : 'wait'}">${c.apiKey ? 'Key saved' : 'Not added yet'}</span></div>
        <p class="muted">Create a key in the <a href="https://platform.claude.com/settings/keys" target="_blank" rel="noopener">Claude Console</a> (API keys). It is stored only in this browser and sent only to Anthropic. Without it, drafts come from your standard answers.</p>
        <div class="row"><input id="wKey" type="password" placeholder="${c.apiKey ? 'Key saved. Paste a new one to replace it' : 'sk-ant-…'}" autocomplete="off" spellcheck="false"><button class="btn" data-act="save-key">Save key</button></div>
      </li>
    </ol>
    <div class="welcome-demo"><span>Waiting for IT? Try everything with sample emails.</span><button class="btn" data-act="demo">${ICON.spark}<span>Open demo</span></button></div>
  </div>`;
}

// ---------- folders & list ----------
async function loadFolders() {
  try {
    S.folders = await S.api.folders();
  } catch (e) {
    fail(e);
  }
  renderSidebar();
}

function allFolders() {
  return [...S.folders.main, ...S.folders.other];
}

function bumpUnread(folderId, delta) {
  const f = allFolders().find((x) => x.id === folderId || x.key === folderId);
  if (f && (f.key === 'inbox' || f.unread)) {
    f.unread = Math.max(0, (f.unread || 0) + delta);
    renderSidebar();
  }
}

function renderSidebar() {
  if (!S.mode) return;
  const cur = S.view === 'mail' && !S.search ? S.folder?.id : null;
  const item = (x) => `<button class="nav-item ${cur === x.id ? 'active' : ''} ${x.child ? 'child' : ''}" data-act="folder" data-id="${esc(x.id)}">${ICON[x.icon] || ICON.folder}<span class="nav-label">${esc(x.name)}</span>${x.unread ? `<span class="count">${x.unread}</span>` : ''}</button>`;
  const other = S.folders.other;
  E.side.innerHTML = `
    <div class="brand">${ICON.logo}<span>Mail Assistant</span>${S.mode === 'demo' ? '<span class="badge">Demo</span>' : ''}</div>
    ${mailboxes().length ? `<label class="mbox" title="Mailbox">${ICON.inbox}<select id="mboxSel" aria-label="Mailbox"><option value="">${esc(S.me?.email || 'My mailbox')}</option>${mailboxes().map((x) => `<option value="${esc(x)}" ${S.mailbox === x ? 'selected' : ''}>${esc(x)}</option>`).join('')}</select></label>` : ''}
    <button class="btn ai block" data-act="compose">${ICON.pen}<span>New email</span></button>
    <nav class="nav">
      <button class="nav-item ${S.view === 'chats' ? 'active' : ''}" data-act="chats">${ICON.chat}<span class="nav-label">Chats</span>${(S.chats || []).reduce((n, c) => n + c.unread, 0) ? `<span class="count">${(S.chats || []).reduce((n, c) => n + c.unread, 0)}</span>` : ''}</button>
      ${S.folders.main.map(item).join('')}
      ${other.length ? `<details class="more" ${other.some((x) => x.id === cur) ? 'open' : ''}><summary>More folders</summary>${other.map(item).join('')}</details>` : ''}
      <div class="nav-sep"></div>
      <button class="nav-item ${S.view === 'contacts' ? 'active' : ''}" data-act="contacts">${ICON.people}<span class="nav-label">Contacts</span></button>
      <button class="nav-item ${S.view === 'settings' ? 'active' : ''}" data-act="settings">${ICON.gear}<span class="nav-label">Settings</span></button>
    </nav>
    <div class="me">${avatar(S.me, 32)}<div class="me-txt"><b>${esc(S.me?.name || '')}</b><small>${esc(S.me?.email || '')}</small></div></div>
    <button class="link small" data-act="${S.mode === 'demo' ? 'exit-demo' : 'signout'}">${S.mode === 'demo' ? 'Exit demo' : 'Sign out'}</button>`;
}

// Chats / Email switch at the top of the list. Chats is where the app starts.
const viewSwitch = (on) => `<div class="view-seg" role="group" aria-label="View">
    <button data-act="chats" class="${on === 'chats' ? 'on' : ''}" aria-pressed="${on === 'chats'}">${ICON.chat}<span>Chats</span></button>
    <button data-act="email" class="${on === 'email' ? 'on' : ''}" aria-pressed="${on === 'email'}">${ICON.mail}<span>Email</span></button>
  </div>`;

function renderListShell() {
  E.list.innerHTML = `
    <div class="list-head">
      <button class="icon-btn only-mobile" data-act="menu" aria-label="Menu">${ICON.menu}</button>
      ${viewSwitch('email')}
      <button class="icon-btn" data-act="refresh" title="Refresh" aria-label="Refresh">${ICON.refresh}</button>
    </div>
    <label class="search">${ICON.search}<input id="q" type="search" placeholder="Search all mail" value="${esc(S.search)}" autocomplete="off"></label>
    <div class="folder-line"><b id="listTitle">${esc(S.search ? 'Search results' : S.folder?.name || '')}</b></div>
    <div class="items" id="items"></div>`;
  renderItems();
}

// Show one row per conversation (newest message), like Outlook's conversation view.
function grouped() {
  const map = new Map();
  for (const m of S.list) {
    const k = m.conversationId || m.id;
    const g = map.get(k);
    if (!g) map.set(k, { ...m, count: 1, anyUnread: !m.isRead });
    else { g.count++; g.anyUnread ||= !m.isRead; }
  }
  return [...map.values()];
}

function renderItems() {
  const box = $('#items');
  if (!box) return;
  if (S.listBusy && !S.list.length) { box.innerHTML = skeleton(7); return; }
  if (!S.list.length) { box.innerHTML = `<div class="empty small">${S.search ? 'No results.' : 'Nothing here.'}</div>`; return; }
  const outgoing = !S.search && /^(sentitems|drafts)$/.test(S.folder?.key || '');
  box.innerHTML = grouped().map((m) => itemHtml(m, outgoing)).join('')
    + (S.next ? '<button class="btn subtle block more-btn" data-act="more">Load more</button>' : '');
}

function itemHtml(m, outgoing) {
  const person = outgoing ? m.to[0] || { name: '(no recipient)' } : m.from;
  const who = outgoing ? `To: ${m.to.map((t) => t.name).join(', ') || '(no recipient)'}` : m.from.name;
  const active = S.sel && (S.sel.id === m.id || (S.sel.conversationId && S.sel.conversationId === m.conversationId));
  return `<div class="item ${m.anyUnread ? 'unread' : ''} ${active ? 'active' : ''}" data-act="open" data-id="${esc(m.id)}" role="button" tabindex="0">
    ${avatar(person)}
    <div class="item-main">
      <div class="item-top"><span class="from">${esc(who)}</span>${m.count > 1 ? `<span class="n">${m.count}</span>` : ''}<time>${esc(fmtDate(m.date))}</time></div>
      <div class="subj">${m.important ? '<span class="imp">!</span>' : ''}${m.isDraft ? '<span class="tag">Draft</span>' : ''}${esc(m.subject)}</div>
      <div class="prev">${esc(m.preview)}</div>
      ${!outgoing && S.chats?.find((c) => c.person.email === m.from?.email)?.last?.id === m.id && hasReadyTag(m) ? READY_TAG : ''}
    </div>
    <div class="item-icons">${m.anyUnread ? '<span class="dot"></span>' : ''}${m.hasAttachments ? ICON.clip : ''}${m.flagged ? `<span class="flagged">${ICON.flag}</span>` : ''}</div>
  </div>`;
}

async function openFolder(f) {
  if (!f) return;
  S.view = 'mail';
  S.folder = f;
  S.search = '';
  S.list = [];
  S.next = null;
  S.listBusy = true;
  const seq = ++S.listSeq;
  renderSidebar();
  renderListShell();
  if (!S.sel) renderEmptyDetail();
  else if (!$('.thread .msgs')) renderThread();
  closeDrawer();
  try {
    const r = await S.api.messages(f.id);
    if (seq !== S.listSeq) return;
    S.list = r.items;
    S.next = r.next;
  } catch (e) {
    fail(e);
  } finally {
    if (seq === S.listSeq) { S.listBusy = false; renderItems(); }
  }
}

const runSearch = debounce(async (q) => {
  S.search = q.trim();
  S.list = [];
  S.next = null;
  S.listBusy = true;
  const seq = ++S.listSeq;
  const title = $('#listTitle');
  if (title) title.textContent = S.search ? 'Search results' : S.folder?.name || '';
  renderSidebar();
  renderItems();
  try {
    const r = await S.api.messages(S.folder?.id, { search: S.search || undefined });
    if (seq !== S.listSeq) return;
    S.list = r.items;
    S.next = r.next;
  } catch (e) {
    fail(e);
  } finally {
    if (seq === S.listSeq) { S.listBusy = false; renderItems(); }
  }
}, 450);

async function loadMore(btn) {
  if (!S.next) return;
  btn.disabled = true;
  btn.textContent = 'Loading…';
  try {
    const r = await S.api.messages(S.folder?.id, { next: S.next });
    S.list.push(...r.items);
    S.next = r.next;
  } catch (e) {
    fail(e);
  }
  renderItems();
}

async function refreshList(silent = false) {
  if (S.view !== 'mail' || !S.folder) return;
  if (S.search) return runSearch(S.search);
  const seq = S.listSeq;
  try {
    const r = await S.api.messages(S.folder.id);
    if (seq !== S.listSeq) return;
    const ids = new Set(r.items.map((m) => m.id));
    S.list = [...r.items, ...S.list.filter((m) => !ids.has(m.id) && new Date(m.date) < new Date(r.items.at(-1)?.date || 0))];
    S.next = S.next || r.next;
    renderItems();
    if (!silent) toast('Up to date', 'ok', 1500);
    S.api.folders().then((f) => { S.folders = f; renderSidebar(); }).catch(() => {});
  } catch (e) {
    if (!silent) fail(e);
  }
}

// ---------- reading ----------
function renderEmptyDetail() {
  E.detail.innerHTML = empty(ICON.mail, 'Pick an email to read', 'Customer emails open with a reply already drafted. Check it and press <b>Send</b>.');
}

async function openMessage(m) {
  S.sel = m;
  S.view = 'mail';
  $$('.item.active').forEach((x) => x.classList.remove('active'));
  $(`.item[data-id="${CSS.escape(m.id)}"]`)?.classList.add('active');
  E.app.classList.add('show-detail');
  E.detail.innerHTML = loading('Opening…');
  try {
    const thread = await S.api.thread(m);
    if (S.sel !== m) return;
    S.thread = thread;
    renderThread();
    const unread = thread.filter((x) => !x.isRead && !x.isDraft);
    if (unread.length) {
      for (const x of unread) {
        x.isRead = true;
        const li = S.list.find((y) => y.id === x.id);
        if (li) li.isRead = true;
        bumpUnread(x.folderId, -1);
        if (can('send')) S.api.markRead(x.id, true).catch(() => {});
      }
      renderItems();
    }
  } catch (e) {
    fail(e);
    if (S.sel === m) E.detail.innerHTML = empty(ICON.mail, "Couldn't open this email", esc(e.message));
  }
}

function renderThread() {
  if (!S.sel) return renderEmptyDetail();
  const t = S.thread;
  const target = replyTarget();
  const st = replyState();
  const open = (m, i) => i === t.length - 1 || !m.isRead || m.id === S.sel.id;
  E.detail.innerHTML = `
  <div class="thread">
    <div class="thread-head">
      <button class="icon-btn only-mobile" data-act="back" aria-label="Back">${ICON.back}</button>
      <h1>${esc(t.at(-1)?.subject?.replace(/^(re|fw|fwd|aw|sv):\s*/i, '') || S.sel.subject)}</h1>
      <div class="thread-actions">
        <button class="btn subtle" data-act="focus-reply">${ICON.spark}<span>Reply</span></button>
        ${S.sel.webLink ? `<a class="icon-btn" href="${esc(S.sel.webLink)}" target="_blank" rel="noopener" title="Open in Outlook" aria-label="Open in Outlook">${ICON.ext}</a>` : ''}
        ${can('send') ? '' : '<!--'}<button class="icon-btn" data-act="mark-unread" title="Mark as unread" aria-label="Mark as unread">${ICON.unread}</button>${can('send') ? '' : '-->'}
      </div>
    </div>
    <div class="msgs">${t.map((m, i) => msgHtml(m, open(m, i))).join('')}</div>
    ${replyPanel(target, st)}
  </div>`;
  t.forEach((m, i) => { if (open(m, i)) mountBody(m); });
  wireDraftArea();
  maybeAutoDraft();
}

function msgHtml(m, open) {
  const mine = m.from?.email && m.from.email === myAddr();
  const to = m.to.map((x) => x.name || x.email).join(', ');
  const cc = m.cc.length ? `, cc ${m.cc.map((x) => x.name || x.email).join(', ')}` : '';
  return `<article class="msg ${open ? 'open' : ''} ${m.isDraft ? 'is-draft' : ''}" data-mid="${esc(m.id)}">
    <header class="msg-head" data-act="toggle-msg">
      ${avatar(m.from)}
      <div class="msg-who">
        <div class="msg-from"><b>${esc(m.from?.name || '')}</b><span class="addr">${esc(m.from?.email || '')}</span>${mine ? '<span class="chip">You</span>' : ''}${m.isDraft ? '<span class="chip wait">Draft</span>' : ''}</div>
        <div class="msg-to">to ${esc(to || '-')}${esc(cc)}</div>
      </div>
      <time title="${esc(fmtFull(m.date))}">${esc(fmtDate(m.date))}</time>
    </header>
    <div class="msg-preview">${esc(m.preview || '')}</div>
    <div class="msg-body">
      <div class="frame-wrap"><iframe sandbox="allow-same-origin allow-popups allow-popups-to-escape-sandbox" title="Email from ${esc(m.from?.name || '')}"></iframe></div>
      <div class="msg-tools"></div>
      <div class="atts"></div>
    </div>
  </article>`;
}

async function mountBody(m, opts = {}) {
  const art = $(`.msg[data-mid="${CSS.escape(m.id)}"]`);
  if (!art) return;
  art.dataset.mounted = '1';
  const frame = $('iframe', art);
  let full = !!opts.full;
  let html = full ? (m.fullHtml ??= await S.api.fullBody(m.id).catch(() => m.bodyHtml)) : m.bodyHtml;
  if (!html?.trim()) {
    html = m.fullHtml ??= await S.api.fullBody(m.id).catch(() => '');
    full = true;
  }
  art.dataset.full = full ? '1' : '';
  if (/cid:/i.test(html)) m.cids ??= await S.api.inlineImages(m.id).catch(() => ({}));
  const allowImages = opts.images ?? (m.imagesAllowed || cfg().loadImages);
  const { html: clean, remoteImages } = cleanEmailHtml(html, m.cids || {});
  frame.style.height = '0px';
  fitFrame(frame);
  frame.srcdoc = srcdocFor(clean || '<p style="color:#98a2b3">(empty message)</p>', allowImages);
  $('.msg-tools', art).innerHTML = [
    full ? (m.bodyHtml && m.fullHtml !== m.bodyHtml ? '<button class="link small" data-act="short-body">Hide quoted text</button>' : '')
      : '<button class="link small" data-act="full-body">Show full message</button>',
    remoteImages && !allowImages ? `<button class="link small" data-act="images">${ICON.eye}<span>Show images</span></button>` : '',
  ].join('');
  if (m.hasAttachments) renderAtts(m, art);
}

function fitFrame(frame) {
  const fit = () => {
    try {
      const d = frame.contentDocument;
      if (!d?.body) return;
      const h = Math.ceil(d.body.getBoundingClientRect().height);
      const scrollbar = d.documentElement.scrollWidth > d.documentElement.clientWidth ? 18 : 0;
      frame.style.height = `${Math.max(24, h + scrollbar + 2)}px`;
    } catch { /* not ready */ }
  };
  frame.onload = () => {
    fit();
    try { new ResizeObserver(fit).observe(frame.contentDocument.body); } catch { /* ignore */ }
    setTimeout(fit, 300);
    setTimeout(fit, 1500);
  };
}

async function renderAtts(m, art) {
  const box = $('.atts', art);
  if (!box || box.dataset.done) return;
  box.dataset.done = '1';
  try {
    m.atts = m.atts?.length ? m.atts : await S.api.attachments(m.id);
    box.innerHTML = m.atts.filter((a) => !a.inline).map((a) => `<button class="att" data-act="att" data-mid="${esc(m.id)}" data-aid="${esc(a.id)}" data-name="${esc(a.name)}">${ICON.clip}<span>${esc(a.name)}</span><small>${esc(fmtSize(a.size))}</small></button>`).join('');
  } catch { /* no attachments shown */ }
}

// ---------- reply: a draft is ready when the email opens ----------
const READY_TAG = `<span class="ready-tag">${ICON.spark}Draft ready</span>`;

function replyTarget() {
  const real = S.thread.filter((m) => !m.isDraft);
  const others = real.filter((m) => m.from?.email && m.from.email !== myAddr());
  return others.at(-1) || real.at(-1) || S.sel;
}

// Already answered when one of my messages came after theirs.
function answered(t) {
  return S.thread.some((m) => !m.isDraft && m.from?.email === myAddr() && new Date(m.date) > new Date(t.date));
}

const aiAuto = () => aiMode() === 'full' && cfg().apiKey && cfg().autoDraft;
const latestText = (t) => (t.text ??= htmlToText((t.bodyHtml || '').replace(/<blockquote[\s\S]*?<\/blockquote>/gi, ''))) || t.preview || '';

// The draft waiting for an email: Claude's (if prepared), otherwise the best standard answer.
function readyFor(t) {
  if (!t || !isCustomer(t.from)) return null;
  const done = S.ready.get(t.id);
  if (done) return { text: done.text, source: 'ai', meta: done.meta || '' };
  if (aiAuto()) return null; // Claude writes it when the chat opens
  const a = pickAnswer(answers(), { subject: t.subject, text: latestText(t) });
  return { text: fillAnswer(a || GENERIC, t.from), source: 'standard', answer: (a || GENERIC).name };
}

// List badge: a prepared Claude draft, or a standard answer that fits.
function hasReadyTag(m) {
  if (!m || m.out || !isCustomer(m.from)) return false;
  if (S.ready.has(m.id)) return true;
  return !aiAuto() && !!pickAnswer(answers(), { subject: m.subject, text: m.preview || '' });
}

function replyState() {
  const key = S.sel.conversationId || S.sel.id;
  if (!S.drafts.has(key)) {
    const t = replyTarget();
    const others = [...t.to, ...t.cc].filter((x) => x.email && x.email !== myAddr() && x.email !== t.from?.email);
    const st = { note: '', text: '', history: null, mode: others.length ? 'all' : 'reply', busy: false, meta: '', source: '', targetId: t.id };
    const r = answered(t) ? null : readyFor(t);
    if (r) Object.assign(st, r);
    S.drafts.set(key, st);
  }
  return S.drafts.get(key);
}

const QUICK = {
  Shorter: 'Make it shorter and tighter.',
  'More formal': 'Make it more formal.',
  Friendlier: 'Make it warmer and friendlier.',
  'Fix grammar only': 'Only fix grammar, spelling and flow. Keep the content and length.',
};

function schedOptions() {
  const now = new Date();
  const at = (days, h) => { const d = new Date(now); d.setDate(d.getDate() + days); d.setHours(h, 0, 0, 0); return d; };
  const mon = at(((8 - now.getDay()) % 7) || 7, 8);
  const list = [['Tomorrow morning', at(1, 8)], ['Tomorrow afternoon', at(1, 13)], ['Monday morning', mon]];
  if (now.getHours() < 12) list.unshift(['This afternoon', at(0, 14)]);
  const seen = new Set();
  return list.filter(([, d]) => !seen.has(+d) && seen.add(+d));
}
const fmtWhen = (d) => new Date(d).toLocaleString('en-GB', { weekday: 'short', day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' });

// Bottom bar: change the draft with Claude, then Schedule or Send (or Reply in Outlook when read-only).
function sendBar(st, kind) {
  const ai = aiMode() !== 'off';
  const isNew = kind === 'new';
  const change = ai
    ? `<label class="change"><span class="sr">Change the draft</span><input id="revise" placeholder="${st.text ? 'Change something… any language' : isNew ? 'What is it about? A few words, any language' : 'What should the reply say? A few words, any language'}" autocomplete="off"><button class="ai-go" data-act="revise" aria-label="Write with AI" ${st.busy ? 'disabled' : ''}>${ICON.spark}</button></label>`
    : '<span class="grow"></span>';
  if (st.pending) {
    return `<div class="as-bar">${change}<div class="pending" role="status"><span class="pending-bar"></span><span>${esc(st.pending.label)}</span><button class="btn" data-act="undo-send">Undo</button></div></div>`;
  }
  if (!canSendHere()) {
    return `<div class="as-bar">${change}<button class="btn primary" data-act="${isNew ? 'outlook-new' : 'outlook-reply'}">${ICON.ext}<span>${isNew ? 'Copy & open Outlook' : 'Reply in Outlook'}</span></button></div>`;
  }
  const menu = st.schedOpen ? `<div class="sched-menu" role="menu" aria-label="Schedule send">
      <div class="sched-h">Send later</div>
      ${schedOptions().map(([l, d]) => `<button role="menuitem" data-act="schedule" data-at="${d.toISOString()}"><span>${esc(l)}</span><small>${esc(fmtWhen(d))}</small></button>`).join('')}
      <div class="sched-pick"><input type="datetime-local" id="schedAt" aria-label="Date and time"><button class="btn" data-act="schedule-pick">Schedule</button></div>
    </div>` : '';
  return `<div class="as-bar">${change}
    <div class="sched-wrap"><button class="btn" data-act="sched-menu" aria-haspopup="menu" aria-expanded="${!!st.schedOpen}">${ICON.clock}<span>Schedule</span>${ICON.chevron}</button>${menu}</div>
    <button class="btn primary" data-act="${isNew ? 'send-new' : 'send-reply'}">${ICON.send}<span>Send</span></button>
  </div>`;
}

function draftLabel(st) {
  if (st.busy) return 'Writing a draft…';
  if (!st.text) return 'Write a reply';
  if (st.source === 'ai') return 'Draft ready';
  if (st.source === 'standard') return st.answer === GENERIC.name ? 'Draft · general reply' : `Draft ready · standard answer: ${st.answer}`;
  return 'Your draft';
}

function replyPanel(target, st) {
  const name = target.from?.name || target.from?.email || '';
  const ai = aiMode() !== 'off';
  const picker = !ai ? `<label class="pick-answer">Standard answer<select id="answerSel"><option value="">Choose…</option>${[...answers(), GENERIC].map((a) => `<option ${a.name === st.answer ? 'selected' : ''}>${esc(a.name)}</option>`).join('')}</select></label>` : '';
  return `<section class="assistant ${st.busy ? 'busy' : ''} ${st.text ? 'has-draft' : ''}" id="assistant">
    <div class="as-head">
      <div class="as-title">${ICON.spark}<b>${esc(draftLabel(st))}</b><span class="muted small">to ${esc(name)}${st.mode === 'all' ? ' and others' : ''}${S.mailbox ? `, from ${esc(S.mailbox)}` : ''}</span></div>
      <div class="seg" role="group" aria-label="Reply type">
        <button data-act="mode" data-mode="reply" class="${st.mode === 'reply' ? 'on' : ''}">Reply</button>
        <button data-act="mode" data-mode="all" class="${st.mode === 'all' ? 'on' : ''}">Reply all</button>
      </div>
    </div>
    <div class="draft ${st.text || st.busy ? '' : 'hidden'}">
      <textarea id="draftText" class="draft-text" spellcheck="true" aria-label="Draft">${esc(st.text)}</textarea>
    </div>
    ${ai && st.text ? `<div class="chips">${Object.keys(QUICK).map((k) => `<button class="chip-btn" data-act="quick-revise" data-rev="${esc(k)}" ${st.busy ? 'disabled' : ''}>${esc(k)}</button>`).join('')}</div>` : picker}
    ${sendBar(st, 'reply')}
    ${st.meta ? `<div class="muted small meta">${esc(st.meta)}</div>` : ''}
    ${!canSendHere() ? '<p class="muted small ro-note">Read-only access: <b>Reply in Outlook</b> copies the draft and opens the email in Outlook. Click Reply there and paste.</p>' : ''}
  </section>`;
}

function rerenderPanel() {
  if (S.view === 'compose') return renderCompose();
  const box = $('#assistant');
  if (!box || !S.sel) return;
  const keep = $('#revise')?.value;
  box.outerHTML = replyPanel(replyTarget(), replyState());
  if (keep && $('#revise')) $('#revise').value = keep;
  wireDraftArea();
}

function wireDraftArea() {
  ['#draftText'].forEach((s) => grow($(s)));
}

// Claude writes the reply as soon as a customer email opens (Full mode with automatic drafts).
function maybeAutoDraft() {
  if (!S.sel || !$('#assistant')) return;
  const st = replyState();
  const t = replyTarget();
  if (st.text || st.busy || st.autoTried || answered(t) || !isCustomer(t.from) || !aiAuto()) return;
  st.autoTried = true;
  runDraft('', { auto: true });
}

function meForAI() {
  const c = cfg();
  return { name: c.myName || S.me?.name || '', role: c.myRole || S.me?.jobTitle || '' };
}

async function chooseModel(c) {
  if (c.model) return c.model;
  const cached = store.get('autoModel');
  if (cached && Date.now() - cached.at < 7 * 864e5) return cached.id;
  try {
    const id = ai.pickModel(await ai.listModels(c.apiKey));
    store.set('autoModel', { id, at: Date.now() });
    return id;
  } catch (e) {
    if (e.status === 401) throw e;
    return ai.FALLBACK_MODEL;
  }
}

async function examples() {
  if (S.examples) return S.examples;
  try { S.examples = await S.api.sentExamples(4, S.me?.email); } catch { S.examples = []; }
  return S.examples;
}

function parseRecipients(s) {
  return String(s || '').split(/[,;]+/).map((part) => {
    const m = part.match(/([^\s<>"]+@[^\s<>"]+\.[^\s<>"]+)/);
    if (!m) return null;
    const name = part.replace(m[0], '').replace(/[<>"]/g, '').trim();
    return { name: name || m[1], email: m[1].toLowerCase() };
  }).filter(Boolean);
}

// Prepared Claude drafts are kept in this browser so they aren't paid for twice.
const readyKey = () => `ready:${S.mailbox || 'me'}`;
function loadReady() {
  S.ready = new Map(Object.entries(store.get(readyKey(), {})));
}
function saveReady(id, v) {
  S.ready.set(id, { ...v, at: Date.now() });
  store.set(readyKey(), Object.fromEntries([...S.ready.entries()].sort((a, b) => b[1].at - a[1].at).slice(0, 40)));
}

// One instruction box does everything: with no draft it writes one from the note, otherwise it revises.
async function runDraft(input, { auto = false } = {}) {
  const isNew = S.view === 'compose';
  const st = isNew ? S.compose : replyState();
  if (!st || st.busy) return;
  const mode = aiMode();
  if (mode === 'off') { toast('AI drafting is off. You can switch it on in Settings.'); return; }
  const c = cfg();
  const v = (input ?? $('#revise')?.value ?? '').trim();
  if (isNew) {
    st.to = $('#cTo')?.value ?? st.to;
    st.subject = $('#cSubject')?.value ?? st.subject;
  }
  const current = ($('#draftText')?.value ?? st.text).trim();
  const target = isNew ? null : replyTarget();
  if (!auto && !v && !(mode === 'full' && !isNew && !current)) return $('#revise')?.focus();

  if (!c.apiKey) {
    if (S.mode === 'demo') {
      const text = current && v ? `${current}\n\n[Demo: Claude would apply “${v}” here.]` : cannedDraft({ target, note: v, isNew });
      if (isNew && !current) { const sp = ai.splitSubject(text); st.subject = sp.subject || st.subject; st.text = sp.body; } else st.text = text;
      st.meta = 'Sample text (demo has no Claude key)';
      st.source = 'ai';
      return rerenderPanel();
    }
    toast('Add your Claude API key in Settings first.', 'err');
    return openSettings('ai');
  }

  const standard = !isNew ? (pickAnswer(answers(), { subject: target.subject, text: latestText(target) })) : null;
  let messages;
  if (current && v) {
    if (mode === 'full' && !isNew) {
      const thread = S.thread.filter((m) => !m.isDraft).map((m) => ({ ...m, text: latestText(m) }));
      messages = [{ role: 'user', content: ai.replyPrompt({ thread, target, note: st.note }) }, { role: 'assistant', content: current }, { role: 'user', content: ai.revisePrompt(v, false) }];
    } else {
      messages = [{ role: 'user', content: ai.reviseDraftPrompt(isNew ? `Subject: ${st.subject}\n\n${current}` : current, v, isNew) }];
    }
  } else if (isNew) {
    messages = [{ role: 'user', content: ai.newEmailPrompt({ to: parseRecipients(st.to), note: v }) }];
  } else if (mode === 'full') {
    st.note = v;
    const thread = S.thread.filter((m) => !m.isDraft).map((m) => ({ ...m, text: latestText(m) }));
    messages = [{ role: 'user', content: ai.replyPrompt({ thread, target, note: v, standard: standard ? fillAnswer(standard, target.from) : '' }) }];
  } else {
    messages = [{ role: 'user', content: ai.noteReplyPrompt({ first: firstName(target.from), note: v, standard: standard ? fillAnswer(standard, target.from) : '' }) }];
  }

  st.busy = true;
  rerenderPanel();
  const t0 = performance.now();
  try {
    const model = await chooseModel(c);
    const system = ai.buildSystem({ me: meForAI(), style: c.style || ai.DEFAULT_STYLE, notes: c.notes, examples: c.useExamples && mode === 'full' ? await examples() : [] });
    const out = await ai.draft({
      key: c.apiKey, model, effort: c.effort, system, messages,
      onText: (t) => {
        const area = $('#draftText');
        if (area && (isNew ? S.view === 'compose' && S.compose === st : S.sel && replyState() === st)) {
          area.closest('.draft')?.classList.remove('hidden');
          area.value = t;
          grow(area);
        }
      },
    });
    if (isNew && !current) {
      const { subject, body } = ai.splitSubject(out.text);
      st.subject = subject || st.subject;
      st.text = body;
    } else if (isNew) {
      const { subject, body } = ai.splitSubject(out.text);
      st.subject = subject || st.subject;
      st.text = body;
    } else {
      st.text = out.text;
    }
    st.source = 'ai';
    st.meta = `${ai.prettyModel(model)} · ${((performance.now() - t0) / 1000).toFixed(1)} s`;
    if (!isNew) saveReady(target.id, { text: st.text, meta: st.meta });
    const inp = $('#revise');
    if (inp) inp.value = '';
  } catch (e) {
    fail(e);
  } finally {
    st.busy = false;
    const visible = isNew ? S.view === 'compose' && S.compose === st : (S.view === 'mail' || S.view === 'chats') && S.sel && replyState() === st;
    if (visible) { rerenderPanel(); if (!isNew) $('#revise') && ($('#revise').value = ''); }
    if (!isNew && S.view === 'chats') renderChatItems();
  }
}

// Background: Claude prepares replies for the newest unanswered customer chats (Full mode only).
async function prepareDrafts() {
  if (!aiAuto() || !S.chats) return;
  const seq = ++S.prepSeq;
  const c = cfg();
  const todo = S.chats
    .filter((x) => x.last && !x.last.out && isCustomer(x.last.from) && !S.ready.has(x.last.id) && Date.now() - new Date(x.last.date) < 14 * 864e5)
    .slice(0, 4);
  for (const chat of todo) {
    if (seq !== S.prepSeq || S.chat?.person.email === chat.person.email) continue;
    try {
      const msgs = await S.api.chat(chat.person.email);
      const target = msgs.filter((m) => m.from?.email !== myAddr()).at(-1);
      if (!target || msgs.at(-1)?.from?.email === myAddr() || S.ready.has(target.id)) continue;
      const thread = msgs.map((m) => ({ ...m, text: latestText(m) }));
      const standard = pickAnswer(answers(), { subject: target.subject, text: latestText(target) });
      const model = await chooseModel(c);
      const out = await ai.draft({
        key: c.apiKey, model, effort: c.effort,
        system: ai.buildSystem({ me: meForAI(), style: c.style || ai.DEFAULT_STYLE, notes: c.notes, examples: c.useExamples ? await examples() : [] }),
        messages: [{ role: 'user', content: ai.replyPrompt({ thread, target, note: '', standard: standard ? fillAnswer(standard, target.from) : '' }) }],
      });
      if (seq !== S.prepSeq) return;
      saveReady(target.id, { text: out.text, meta: `${ai.prettyModel(model)} · prepared` });
      if (S.view === 'chats') renderChatItems();
    } catch (e) {
      if (e.status === 401) return fail(e);
    }
  }
}

function signatureHtml(c) {
  return c.addSignature && c.signatureHtml ? `<div>${sanitizeBasic(c.signatureHtml)}</div>` : '';
}

// ---------- sending: Send waits 5 seconds so it can be undone ----------
function queueSend(isNew, at = null) {
  const st = isNew ? S.compose : replyState();
  const text = $('#draftText')?.value.trim();
  if (!text) return toast('Nothing to send yet.');
  if (/\[[^\]\n]{2,60}\]/.test(text) && !st.bracketOk) {
    st.bracketOk = true;
    return toast('The draft still has [placeholders]. Fill them in, or press again to send anyway.', 'err', 6000);
  }
  st.text = text;
  st.schedOpen = false;
  if (at && new Date(at) < Date.now() + 60e3) return toast('Pick a time in the future.', 'err');
  const c = cfg();
  const ctx = { st, isNew, at, text, html: textToHtml(text, c.draftFont) + signatureHtml(c), mailbox: S.mailbox };
  let who;
  if (isNew) {
    st.to = $('#cTo').value;
    st.subject = $('#cSubject').value.trim();
    ctx.to = parseRecipients(st.to);
    ctx.subject = st.subject;
    if (!ctx.to.length) return toast('Add at least one recipient email address.', 'err');
    if (!ctx.subject) return toast('Add a subject first.', 'err');
    who = ctx.to.map((x) => x.name.split(' ')[0]).slice(0, 2).join(', ');
  } else {
    ctx.target = replyTarget();
    ctx.replyAll = st.mode === 'all';
    ctx.chatEmail = S.view === 'chats' ? S.chat?.person.email : null;
    who = firstName(ctx.target.from) || ctx.target.from?.email;
  }
  st.pending = { label: at ? `Scheduling for ${fmtWhen(at)}…` : `Sending to ${who}…`, timer: setTimeout(() => doSend(ctx), 5000) };
  rerenderPanel();
}

function undoSend() {
  const st = S.view === 'compose' ? S.compose : S.sel && replyState();
  if (!st?.pending) return;
  clearTimeout(st.pending.timer);
  st.pending = null;
  rerenderPanel();
  toast('Not sent', 'ok', 1500);
}

async function doSend(ctx) {
  const { st } = ctx;
  const here = () => (ctx.isNew ? S.view === 'compose' && S.compose === st : S.sel && S.drafts.get(S.sel.conversationId || S.sel.id) === st);
  if (ctx.mailbox !== S.mailbox) { st.pending = null; return toast('Not sent: the mailbox changed.', 'err'); }
  try {
    const draft = ctx.isNew
      ? await S.api.createDraft({ to: ctx.to, subject: ctx.subject, html: ctx.html })
      : await S.api.createReplyDraft(ctx.target.id, { replyAll: ctx.replyAll, html: ctx.html });
    await S.api.send(draft.id, { at: ctx.at });
    toast(ctx.at ? `Scheduled for ${fmtWhen(ctx.at)}. It waits in your Outbox until then.` : `Sent${ctx.mailbox ? ` from ${ctx.mailbox}` : ''}`, 'ok', 4500);
  } catch (e) {
    st.pending = null;
    fail(e);
    if (here()) rerenderPanel();
    return;
  }
  st.pending = null;
  if (ctx.isNew) {
    S.compose = { to: '', subject: '', note: '', text: '', busy: false, meta: '', history: null };
    if (S.view === 'compose') renderCompose();
    return;
  }
  // Show the reply straight away (Outlook's search index can take a minute to catch up).
  const t = ctx.target;
  const sent = {
    id: `sent-${Date.now()}`, conversationId: t.conversationId, from: { name: S.mailbox || S.me?.name || 'Me', email: myAddr() },
    to: [t.from], cc: [], date: new Date().toISOString(), subject: `RE: ${t.subject.replace(/^re:\s*/i, '')}`,
    bodyHtml: ctx.html, preview: ctx.text.slice(0, 140), isRead: true, isDraft: false, hasAttachments: false, scheduled: ctx.at || null,
  };
  Object.assign(st, { note: '', text: '', history: null, meta: '', source: '', answer: '', autoTried: true, bracketOk: false });
  S.ready.delete(t.id);
  const entry = S.chats?.find((x) => x.person.email === (ctx.chatEmail || t.from?.email));
  if (entry) entry.last = { ...sent, out: true };
  if (S.view === 'chats' && S.chat?.msgs && S.chat.person.email === ctx.chatEmail) {
    S.chat.msgs.push(sent);
    drawChat();
  } else if (S.view === 'mail' && here()) {
    S.thread.push(sent);
    renderThread();
  }
  if (S.view === 'chats') renderChatItems();
}

// Read-only access: copy the draft, open the email in Outlook, paste into Outlook's reply.
async function replyInOutlook(isNew) {
  const text = $('#draftText')?.value.trim() || '';
  const ok = text ? await copyRich(text, textToHtml(text, cfg().draftFont)) : true;
  let url;
  if (isNew) {
    const to = parseRecipients($('#cTo')?.value).map((x) => x.email).join(',');
    url = `https://outlook.office.com/mail/deeplink/compose?to=${encodeURIComponent(to)}&subject=${encodeURIComponent($('#cSubject')?.value || '')}`;
  } else url = replyTarget().webLink;
  if (url) window.open(url, '_blank', 'noopener');
  toast(!url ? 'Copied. (Demo emails have no Outlook link.)' : ok ? (isNew ? 'Draft copied. Paste it into the new email in Outlook.' : 'Draft copied. In Outlook, click Reply and paste.') : 'Opened Outlook. Copy failed: select the draft and press Ctrl+C.', ok ? 'ok' : 'err', 6000);
}

// ---------- new email ----------
function openCompose(prefill = {}) {
  S.view = 'compose';
  S.compose = { to: prefill.to || '', subject: '', note: '', text: '', busy: false, meta: '', history: null };
  renderSidebar();
  closeDrawer();
  E.app.classList.add('show-detail');
  if (!$('#items') && !$('#chitems')) renderListShell();
  renderCompose();
  $(prefill.to ? '#cSubject' : '#cTo')?.focus();
}

function renderCompose() {
  const c = S.compose;
  E.detail.innerHTML = `
  <div class="thread compose">
    <div class="thread-head"><button class="icon-btn only-mobile" data-act="back" aria-label="Back">${ICON.back}</button><h1>New email</h1></div>
    <section class="assistant ${c.busy ? 'busy' : ''}" id="assistant">
      <label class="field">To<input id="cTo" value="${esc(c.to)}" placeholder="Name or email. Start typing to search" autocomplete="off"></label>
      <div class="suggest" id="cSuggest"></div>
      <label class="field">Subject<input id="cSubject" value="${esc(c.subject)}" autocomplete="off"></label>
      <div class="draft"><textarea id="draftText" class="draft-text" spellcheck="true" aria-label="Email text" placeholder="${aiMode() !== 'off' ? 'Write here, or describe the email below and Claude writes it.' : 'Write your email here.'}">${esc(c.text)}</textarea></div>
      ${aiMode() !== 'off' && c.text ? `<div class="chips">${Object.keys(QUICK).map((k) => `<button class="chip-btn" data-act="quick-revise" data-rev="${esc(k)}">${esc(k)}</button>`).join('')}</div>` : ''}
      ${sendBar(c, 'new')}
      ${c.meta ? `<div class="muted small meta">${esc(c.meta)}</div>` : ''}
    </section>
  </div>`;
  wireDraftArea();
}

const suggest = debounce(async (value) => {
  const box = $('#cSuggest');
  if (!box) return;
  const token = value.split(/[,;]/).pop().trim();
  if (token.length < 2 || token.includes('@')) { box.innerHTML = ''; return; }
  try {
    const people = (await findPeople(token)).slice(0, 6);
    if (S.view !== 'compose') return;
    box.innerHTML = people.map((x) => `<button class="sugg" data-act="pick" data-name="${esc(x.name)}" data-email="${esc(x.email)}">${avatar(x, 24)}<span>${esc(x.name)}</span><small>${esc(x.email)}</small></button>`).join('');
  } catch { box.innerHTML = ''; }
}, 300);

// People search: Outlook's people list when allowed, otherwise the people in your chats.
async function findPeople(q = '') {
  if (can('people')) return S.api.people(q);
  if (!S.chats) S.chats = await S.api.chats(myAddr()).catch(() => []);
  const ql = q.toLowerCase();
  return S.chats.map((c) => ({ ...c.person, company: '', jobTitle: '', phones: [] })).filter((x) => !ql || `${x.name} ${x.email}`.toLowerCase().includes(ql));
}

// ---------- chats: all emails with one person, messenger style ----------
async function openChats() {
  S.view = 'chats';
  renderSidebar();
  closeDrawer();
  E.list.innerHTML = `
    <div class="list-head">
      <button class="icon-btn only-mobile" data-act="menu" aria-label="Menu">${ICON.menu}</button>
      ${viewSwitch('chats')}
      <button class="icon-btn" data-act="chats-refresh" title="Refresh" aria-label="Refresh">${ICON.refresh}</button>
    </div>
    <label class="search">${ICON.search}<input id="chq" type="search" placeholder="Search people" value="${esc(S.chatQ)}" autocomplete="off"></label>
    <div class="items" id="chitems">${S.chats ? '' : skeleton(7)}</div>`;
  if (S.chat) drawChat();
  else E.detail.innerHTML = empty(ICON.chat, 'Pick a chat', 'Every email with one person in a single conversation, like a messenger. Reply with a short note at the bottom.');
  if (S.chats) renderChatItems();
  try {
    S.chats = await S.api.chats(myAddr());
  } catch (e) {
    fail(e);
    S.chats ||= [];
  }
  if (S.view === 'chats') { renderChatItems(); renderSidebar(); }
  prepareDrafts();
}

function renderChatItems() {
  const box = $('#chitems');
  if (!box) return;
  const q = S.chatQ.trim().toLowerCase();
  const match = (c) => !q || `${c.person.name} ${c.person.email}`.toLowerCase().includes(q);
  const chats = S.chats || [];
  const pins = getPins();
  const pinned = pins.map((p) => chats.find((c) => c.person.email === p.email) || { person: p, last: null, unread: 0 }).filter(match);
  const rest = chats.filter((c) => !isPinned(c.person.email)).filter(match);
  const row = (c, pin) => `<div class="item chat-item ${c.unread ? 'unread' : ''} ${S.chat?.person.email === c.person.email ? 'active' : ''}" data-act="chat" data-email="${esc(c.person.email)}" role="button" tabindex="0">
      ${avatar(c.person, 44)}
      <div class="item-main">
        <div class="item-top"><span class="from">${pin ? `<span class="pin-ic">${ICON.pin}</span>` : ''}${esc(c.person.name)}</span>${c.last ? `<time>${esc(fmtDate(c.last.date))}</time>` : ''}</div>
        <div class="prev">${c.last ? `${c.last.out ? `<span class="you">${c.last.scheduled ? 'Scheduled: ' : 'You: '}</span>` : ''}${esc(c.last.preview || c.last.subject)}` : `<span class="you">${esc(c.person.email)}</span>`}</div>
        ${hasReadyTag(c.last) ? READY_TAG : ''}
      </div>
      ${c.unread ? `<span class="badge-n">${c.unread}</span>` : ''}
    </div>`;
  const head = (t) => (pinned.length ? `<div class="list-section">${t}</div>` : '');
  box.innerHTML = pinned.length || rest.length
    ? (pinned.length ? head('Pinned') + pinned.map((c) => row(c, true)).join('') : '') + (rest.length ? head('Recent') + rest.map((c) => row(c, false)).join('') : '')
    : `<div class="empty small">${q ? 'No matches.' : 'No conversations yet.'}</div>`;
}

async function openChat(person) {
  const seq = ++S.chatSeq;
  S.chat = { person, msgs: null };
  S.view = 'chats';
  $$('#chitems .item').forEach((el) => el.classList.toggle('active', el.dataset.email === person.email));
  E.app.classList.add('show-detail');
  E.detail.innerHTML = loading('Opening chat…');
  let msgs;
  try {
    msgs = await S.api.chat(person.email);
  } catch (e) {
    fail(e);
    E.detail.innerHTML = empty(ICON.chat, "Couldn't open this chat", esc(e.message));
    return;
  }
  if (seq !== S.chatSeq) return;
  S.chat.msgs = msgs;
  drawChat();
  const unread = msgs.filter((m) => !m.isRead && m.from?.email !== myAddr());
  for (const m of unread) {
    m.isRead = true;
    bumpUnread(m.folderId, -1);
    if (can('send')) S.api.markRead(m.id, true).catch(() => {});
  }
  const entry = S.chats?.find((c) => c.person.email === person.email);
  if (entry && entry.unread) { entry.unread = 0; renderChatItems(); renderSidebar(); }
}

function drawChat() {
  const { person, msgs } = S.chat;
  if (!msgs) return;
  const incoming = msgs.filter((m) => m.from?.email !== myAddr());
  // Drafting replies to the latest email they sent; the whole chat is the context.
  S.sel = incoming.at(-1) || msgs.at(-1) || null;
  S.thread = msgs;
  let day = '', subj = '';
  const norm = (s) => (s || '').replace(/^((re|fw|fwd|aw|sv):\s*)+/i, '').trim().toLowerCase();
  const bubbles = msgs.map((m) => {
    const out = m.from?.email === myAddr();
    const text = (m.chatText ??= htmlToText((m.bodyHtml || '').replace(/<blockquote[\s\S]*?<\/blockquote>/gi, '')));
    const short = text.length > 700 ? `${text.slice(0, 700).trim()}…` : text;
    const d = dayLabel(m.date);
    const showSubj = norm(m.subject) !== subj;
    subj = norm(m.subject);
    const sep = d !== day ? `<div class="chat-day">${esc(d)}</div>` : '';
    day = d;
    return `${sep}<div class="b-row ${out ? 'out' : 'in'}">
      <div class="bubble">
        ${showSubj ? `<div class="b-subj">${esc(m.subject.replace(/^((re|fw|fwd|aw|sv):\s*)+/i, ''))}</div>` : ''}
        <div class="b-text">${esc(short || m.preview || '(empty message)')}</div>
        <div class="b-meta">${m.hasAttachments ? `<span class="b-att">${ICON.clip}</span>` : ''}${m.cc?.length ? `<span>+${m.cc.length} cc</span>` : ''}${m.scheduled ? `<span class="sched">${ICON.clock}Scheduled ${esc(fmtWhen(m.scheduled))}</span>` : `<button class="link" data-act="chat-open" data-id="${esc(m.id)}">Open email</button>`}<time title="${esc(fmtFull(m.date))}">${esc(new Date(m.date).toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' }))}</time></div>
      </div>
    </div>`;
  }).join('');
  E.detail.innerHTML = `
  <div class="chat">
    <div class="chat-head">
      <button class="icon-btn only-mobile" data-act="back" aria-label="Back">${ICON.back}</button>
      ${avatar(person, 40)}
      <div class="grow"><b>${esc(person.name)}</b><small>${esc(person.email)}</small></div>
      <button class="icon-btn ${isPinned(person.email) ? 'on' : ''}" data-act="pin" data-email="${esc(person.email)}" title="${isPinned(person.email) ? 'Unpin' : 'Pin to top'}" aria-label="${isPinned(person.email) ? 'Unpin' : 'Pin to top'}">${ICON.pin}</button>
      <button class="icon-btn" data-act="chat-contact" title="Contact details" aria-label="Contact details">${ICON.people}</button>
    </div>
    <div class="chat-body" id="chatBody">${bubbles || '<div class="empty small">No emails with this person yet.</div>'}</div>
    ${S.sel ? replyPanel(replyTarget(), replyState()) : ''}
  </div>`;
  const body = $('#chatBody');
  body.scrollTop = body.scrollHeight;
  wireDraftArea();
  maybeAutoDraft();
}

// ---------- contacts ----------
function openContacts() {
  S.view = 'contacts';
  renderSidebar();
  closeDrawer();
  E.list.innerHTML = `
    <div class="list-head"><button class="icon-btn only-mobile" data-act="menu" aria-label="Menu">${ICON.menu}</button><h2>Contacts</h2></div>
    <label class="search">${ICON.search}<input id="cq" type="search" placeholder="Search people" value="${esc(S.contactQ)}" autocomplete="off"></label>
    <div class="tabs">
      <button data-act="ctab" data-tab="frequent" class="${S.contactsTab === 'frequent' ? 'on' : ''}">People you email</button>
      ${can('contacts') ? `<button data-act="ctab" data-tab="outlook" class="${S.contactsTab === 'outlook' ? 'on' : ''}">Saved contacts</button>` : ''}
    </div>
    <div class="items" id="citems">${skeleton(6)}</div>`;
  if (S.contact) renderContact(S.contact);
  else E.detail.innerHTML = empty(ICON.people, 'Pick a contact', 'See their details and your recent emails with them.');
  loadContacts();
}

async function loadContacts() {
  const q = S.contactQ.trim().toLowerCase();
  const seq = ++S.contactSeq;
  let list = [];
  try {
    if (S.contactsTab === 'outlook' && can('contacts')) {
      S.contacts ??= await S.api.contacts();
      list = q ? S.contacts.filter((x) => `${x.name} ${x.email} ${x.company}`.toLowerCase().includes(q)) : S.contacts;
    } else {
      list = await findPeople(q);
    }
  } catch (e) {
    fail(e);
  }
  if (seq !== S.contactSeq || S.view !== 'contacts') return;
  S.people = list;
  const box = $('#citems');
  if (!box) return;
  box.innerHTML = list.length
    ? list.map((x, i) => `<div class="item person ${S.contact?.email === x.email ? 'active' : ''}" data-act="contact" data-i="${i}" role="button" tabindex="0">${avatar(x)}<div class="item-main"><div class="from">${esc(x.name)}</div><div class="prev">${esc([x.company, x.email].filter(Boolean).join(' · '))}</div></div></div>`).join('')
    : `<div class="empty small">${q ? 'No matches.' : 'No contacts found.'}</div>`;
}

async function renderContact(x) {
  S.contact = x;
  E.app.classList.add('show-detail');
  $$('#citems .item').forEach((el) => el.classList.toggle('active', S.people[el.dataset.i]?.email === x.email));
  const emails = (x.emails?.length ? x.emails : [x.email]).filter(Boolean);
  E.detail.innerHTML = `
  <div class="thread contact">
    <div class="thread-head"><button class="icon-btn only-mobile" data-act="back" aria-label="Back">${ICON.back}</button></div>
    <div class="card contact-card">
      ${avatar(x, 64)}
      <div class="grow"><h1>${esc(x.name)}</h1>${x.jobTitle || x.company ? `<div class="muted">${esc([x.jobTitle, x.company].filter(Boolean).join(' · '))}</div>` : ''}</div>
      <button class="btn" data-act="pin" data-email="${esc(x.email)}" data-name="${esc(x.name)}">${ICON.pin}<span>${isPinned(x.email) ? 'Unpin' : 'Pin to Chats'}</span></button>
      <button class="btn ai" data-act="write-to">${ICON.pen}<span>Write email</span></button>
    </div>
    <div class="card kv">
      ${emails.map((e) => `<div class="kv-row">${ICON.mail}<a href="mailto:${esc(e)}">${esc(e)}</a><button class="link small" data-act="copy-text" data-text="${esc(e)}">Copy</button></div>`).join('')}
      ${(x.phones || []).map((ph) => `<div class="kv-row">${ICON.phone}<a href="tel:${esc(ph.replace(/[^\d+]/g, ''))}">${esc(ph)}</a><button class="link small" data-act="copy-text" data-text="${esc(ph)}">Copy</button></div>`).join('')}
      ${x.company ? `<div class="kv-row">${ICON.building}<span>${esc(x.company)}</span></div>` : ''}
    </div>
    <h3 class="section-h">Recent emails</h3>
    <div class="card list-card" id="withList">${skeleton(3)}</div>
  </div>`;
  if (!x.email) { $('#withList').innerHTML = '<div class="empty small">No email address.</div>'; return; }
  try {
    const r = await S.api.messages(null, { search: `participants:${x.email}` });
    if (S.contact !== x) return;
    S.withList = r.items;
    $('#withList').innerHTML = r.items.length
      ? r.items.slice(0, 15).map((m, i) => `<div class="item compact" data-act="open-with" data-i="${i}" role="button" tabindex="0"><div class="item-main"><div class="item-top"><span class="from">${esc(m.from.email === myAddr() ? `You → ${m.to.map((t) => t.name).join(', ')}` : m.from.name)}</span><time>${esc(fmtDate(m.date))}</time></div><div class="subj">${esc(m.subject)}</div><div class="prev">${esc(m.preview)}</div></div></div>`).join('')
      : '<div class="empty small">No emails found.</div>';
  } catch {
    if (S.contact === x) $('#withList').innerHTML = '<div class="empty small">Couldn\'t load emails.</div>';
  }
}

// ---------- settings ----------
function openSettings(section) {
  S.view = 'settings';
  renderSidebar();
  closeDrawer();
  E.app.classList.add('show-detail');
  const c = cfg();
  E.detail.innerHTML = `
  <div class="thread settings">
    <div class="thread-head"><button class="icon-btn only-mobile" data-act="back" aria-label="Back">${ICON.back}</button><h1>Settings</h1><span class="saved-note" id="savedNote"></span></div>

    <section class="card" id="sec-ms"><h3>Microsoft 365</h3>
      ${S.mode === 'live' ? `<p>Signed in as <b>${esc(S.me?.email || '')}</b>.</p>` : S.mode === 'demo' ? '<p class="muted">You are in demo mode, which shows every feature. Add the IDs from IT, then use <b>Exit demo</b> in the sidebar and sign in.</p>' : ''}
      <div class="grid2">
        <label>Application (client) ID<input data-set="clientId" value="${esc(c.clientId)}" placeholder="${esc(CONFIG.clientId || 'from IT')}" spellcheck="false" autocomplete="off"></label>
        <label>Directory (tenant) ID<input data-set="tenantId" value="${esc(c.tenantId)}" placeholder="${esc(CONFIG.tenantId || 'from IT')}" spellcheck="false" autocomplete="off"></label>
      </div>
      <p class="muted small">Redirect URI for IT: <code>${esc(auth.redirectUri())}</code> <button class="link" data-act="copy-text" data-text="${esc(auth.redirectUri())}">Copy</button></p>
      <h4>What the app may do in Outlook</h4>
      <p class="muted small">It starts read-only. Switch on only what IT has approved; Microsoft asks for the extra permission the next time you sign in.</p>
      <div class="perm"><span class="check on">${ICON.shield}</span><div class="grow"><b>Read my mail</b><small>User.Read, Mail.Read · always on</small></div>${permChip(true, true)}</div>
      ${permRow('accSend', 'Send and schedule from the app', 'Mail.Send, Mail.ReadWrite. Without it: Reply in Outlook (copy and paste).', c.accSend, can('send'))}
      ${permRow('accShared', 'Shared mailboxes', `${c.accSend ? 'Mail.ReadWrite.Shared, Mail.Send.Shared' : 'Mail.Read.Shared'}. Mailboxes you can already open in Outlook.`, c.accShared, can('shared'))}
      ${permRow('accContacts', 'Outlook contacts and people search', 'Contacts.Read, People.Read. Without it: people come from your chats.', c.accContacts, can('contacts'))}
      ${S.mode === 'live' && ((c.accSend && !can('send')) || (c.accShared && !can('shared')) || (c.accContacts && !can('contacts'))) ? '<div class="row"><button class="btn primary" data-act="signin-again">Sign in again to allow</button><span class="muted small">If Microsoft says admin approval is needed, IT has to approve it first.</span></div>' : ''}
      ${c.accShared || S.mode === 'demo' ? `<label>Shared mailboxes <span class="muted small">(one per line, e.g. info@company.com)</span>
        <textarea data-set="sharedMailboxes" rows="3" placeholder="info@company.com" spellcheck="false">${esc(c.sharedMailboxes)}</textarea></label>` : ''}
    </section>

    <section class="card" id="sec-ai"><h3>Claude (drafting)</h3>
      <div class="radios" role="radiogroup" aria-label="What Claude may see">
        ${[
          ['off', 'Off', 'Nothing leaves Microsoft and this browser. Drafts come from your standard answers; you edit them yourself.'],
          ['note', 'My draft only', "Claude sees your draft and what you type, never the customer's email."],
          ['full', 'Full', 'Claude reads the email thread and writes the reply. Best drafts.'],
        ].map(([v, l, d]) => `<label class="radio ${c.aiMode === v ? 'on' : ''}"><input type="radio" name="aiMode" data-set="aiMode" value="${v}" ${c.aiMode === v ? 'checked' : ''}><span><b>${l}</b><small>${d}</small></span></label>`).join('')}
      </div>
      <label class="check"><input type="checkbox" data-set="autoDraft" ${c.autoDraft ? 'checked' : ''} ${c.aiMode !== 'full' ? 'disabled' : ''}> Prepare replies automatically for new customer emails (Full mode, about 1 to 3 US cents each)</label>
      <label>API key
        <div class="row"><input id="apiKey" type="password" value="${esc(c.apiKey)}" placeholder="sk-ant-…" autocomplete="off" spellcheck="false"><button class="btn" data-act="toggle-key">Show</button><button class="btn" data-act="test-key">Test</button></div>
      </label>
      <p class="muted small">Stored only in this browser and sent only to Anthropic. Create one in the <a href="https://platform.claude.com/settings/keys" target="_blank" rel="noopener">Claude Console</a>.</p>
      <div class="grid2">
        <label>Model<select data-set="model" id="modelSel"><option value="">Automatic (newest Sonnet)</option>${c.model ? `<option value="${esc(c.model)}" selected>${esc(ai.prettyModel(c.model))}</option>` : ''}</select></label>
        <label>Speed<select data-set="effort">${[['low', 'Fast'], ['medium', 'Balanced'], ['high', 'Thorough']].map(([v, l]) => `<option value="${v}" ${c.effort === v ? 'selected' : ''}>${l}</option>`).join('')}</select></label>
      </div>
    </section>

    <section class="card" id="sec-answers"><h3>Standard answers</h3>
      <p class="muted small">The app picks the best fit by keywords and has it ready when the email opens. Start each answer with <code>## Name</code>, then <code>keywords: …</code>. <code>{first}</code> becomes the first name; <code>[brackets]</code> mark what to fill in (Send warns you if any are left).</p>
      <textarea data-set="answers" rows="16" spellcheck="false">${esc(c.answers || DEFAULT_ANSWERS)}</textarea>
      <div class="row end"><button class="link small" data-act="reset-answers">Reset to default</button></div>
    </section>

    <section class="card" id="sec-style"><h3>Writing style</h3>
      <div class="grid2">
        <label>Your name<input data-set="myName" value="${esc(c.myName)}" placeholder="${esc(S.me?.name || 'Your name')}" autocomplete="off"></label>
        <label>Your role<input data-set="myRole" value="${esc(c.myRole)}" placeholder="${esc(S.me?.jobTitle || 'e.g. Sales Manager')}" autocomplete="off"></label>
      </div>
      <label>Style guide<textarea data-set="style" rows="14">${esc(c.style || ai.DEFAULT_STYLE)}</textarea></label>
      <div class="row end"><button class="link small" data-act="reset-style">Reset to default</button></div>
      <label>Private notes and templates <span class="muted small">(stays in this browser)</span>
        <textarea data-set="notes" rows="10" placeholder="Standard answers, prices, lead times, policies… Claude uses them only when relevant.">${esc(c.notes)}</textarea></label>
      <label class="check"><input type="checkbox" data-set="useExamples" ${c.useExamples ? 'checked' : ''}> Learn my tone from my 4 most recent sent emails to people outside the company</label>
    </section>

    <section class="card" id="sec-out"><h3>Sending</h3>
      <label class="check"><input type="checkbox" data-set="addSignature" ${c.addSignature ? 'checked' : ''}> Add my signature to emails sent from the app</label>
      <p class="muted small">Outlook doesn't add your signature to emails sent by apps, so paste it below once (copy it from one of your sent emails). With <b>Reply in Outlook</b>, Outlook adds it as usual.</p>
      <div class="sig" id="sigBox" contenteditable="true" aria-label="Signature">${sanitizeBasic(c.signatureHtml || '')}</div>
      <div class="grid2">
        <label>Email font<input data-set="draftFont" value="${esc(c.draftFont)}" autocomplete="off"></label>
        <label class="check pad"><input type="checkbox" data-set="loadImages" ${c.loadImages ? 'checked' : ''}> Load images in emails automatically</label>
      </div>
    </section>

    <section class="card"><h3>Backup</h3>
      <p class="muted small">Move your style, notes and signature to another browser. The API key is not included.</p>
      <div class="row"><button class="btn" data-act="export">Export settings</button><label class="btn">Import settings<input type="file" id="importFile" accept="application/json,.json" hidden></label></div>
    </section>
  </div>`;
  $$('.settings textarea').forEach(grow);
  if (c.apiKey) fillModels();
  if (section) $(`#sec-${section}`)?.scrollIntoView({ block: 'start' });
}

const permChip = (want, has) => (S.mode !== 'live' ? '' : has ? '<span class="chip ok">Allowed</span>' : want ? '<span class="chip wait">Sign in to allow</span>' : '');
const permRow = (key, title, sub, on, has) => `<label class="perm"><input type="checkbox" data-set="${key}" ${on ? 'checked' : ''}><div class="grow"><b>${esc(title)}</b><small>${esc(sub)}</small></div>${permChip(on, has)}</label>`;

async function fillModels() {
  const sel = $('#modelSel');
  const c = cfg();
  if (!sel || !c.apiKey) return;
  try {
    const models = await ai.listModels(c.apiKey);
    sel.innerHTML = `<option value="">Automatic (${esc(ai.prettyModel(ai.pickModel(models)))})</option>`
      + models.map((m) => `<option value="${esc(m.id)}" ${m.id === c.model ? 'selected' : ''}>${esc(m.display_name || ai.prettyModel(m.id))}</option>`).join('');
    return models;
  } catch (e) {
    return null;
  }
}

const flashSaved = debounce(() => {
  const n = $('#savedNote');
  if (!n) return;
  n.textContent = 'Saved';
  n.classList.add('on');
  setTimeout(() => n.classList.remove('on'), 1200);
}, 150);
const pendingSaves = {};
function saveField(key, value) {
  clearTimeout(pendingSaves[key]);
  pendingSaves[key] = setTimeout(() => {
    setCfg({ [key]: value });
    flashSaved();
    if (key === 'sharedMailboxes') renderSidebar();
    if (key === 'answers') S.drafts = new Map();
  }, 350);
}
function keepScroll(fn) {
  const y = E.detail.scrollTop;
  fn();
  E.detail.scrollTop = y;
}

// ---------- layout helpers ----------
const loading = (t) => `<div class="loading"><span class="spinner"></span>${esc(t)}</div>`;
const empty = (icon, title, sub = '') => `<div class="empty big"><div class="empty-ic">${icon}</div><b>${esc(title)}</b><p>${sub}</p></div>`;
const skeleton = (k) => Array.from({ length: k }, () => '<div class="item skel"><span class="av"></span><div class="item-main"><i></i><i></i><i></i></div></div>').join('');
function closeDrawer() { E.app.classList.remove('drawer'); }

// ---------- events ----------
function wireEvents() {
  document.addEventListener('click', (ev) => onClick(ev).catch(fail));
  document.addEventListener('input', onInput);
  document.addEventListener('change', onChange);
  document.addEventListener('keydown', onKey);
  setInterval(() => {
    if (S.mode && S.view === 'mail' && !S.search && document.visibilityState === 'visible') refreshList(true);
  }, 120000);
}

async function onClick(ev) {
  const t = ev.target.closest('[data-act]');
  if (!t) {
    if (E.app.classList.contains('drawer') && !ev.target.closest('.sidebar')) closeDrawer();
    return;
  }
  const act = t.dataset.act;
  switch (act) {
    case 'signin': {
      const clientId = ($('#wClient')?.value || '').trim();
      const tenantId = ($('#wTenant')?.value || '').trim();
      if ($('#wClient')) setCfg({ clientId, tenantId });
      const cfgNow = auth.msConfig();
      if (!cfgNow.clientId) return toast('Paste the Application (client) ID from IT first.', 'err');
      if (!GUID.test(cfgNow.clientId)) return toast('The client ID should look like xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx.', 'err');
      if (!tenantId && !CONFIG.tenantId) return toast('Paste the Directory (tenant) ID from IT too.', 'err');
      return auth.login();
    }
    case 'save-key': {
      const v = $('#wKey').value.trim();
      if (!v) return toast('Paste your API key first.', 'err');
      setCfg({ apiKey: v });
      store.del('autoModel');
      toast('API key saved in this browser.', 'ok');
      return renderWelcome();
    }
    case 'demo': return startDemo();
    case 'exit-demo': store.del('demo'); return renderWelcome();
    case 'signout': auth.signOut(); return renderWelcome();
    case 'copy-text': {
      const ok = await copyRich(t.dataset.text);
      return toast(ok ? 'Copied' : 'Copy failed', ok ? 'ok' : 'err', 1500);
    }
    case 'menu': return E.app.classList.toggle('drawer');
    case 'back':
      E.app.classList.remove('show-detail');
      if (S.view === 'compose' || S.view === 'settings') { S.view = 'mail'; renderSidebar(); renderListShell(); }
      return;
    case 'folder': {
      const f = allFolders().find((x) => x.id === t.dataset.id);
      S.sel = null;
      E.app.classList.remove('show-detail');
      return openFolder(f);
    }
    case 'chats': return openChats();
    case 'email': {
      S.sel = null;
      S.chat = null;
      if (!S.chats) S.api.chats(myAddr()).then((c) => { S.chats = c; renderItems(); }).catch(() => {});
      E.app.classList.remove('show-detail');
      return openFolder(S.folder || S.folders.main[0] || { id: 'inbox', name: 'Inbox', key: 'inbox' });
    }
    case 'chats-refresh': S.chats = null; return openChats();
    case 'chat': {
      const p = S.chats?.find((x) => x.person.email === t.dataset.email)?.person || getPins().find((x) => x.email === t.dataset.email);
      return p && openChat(p);
    }
    case 'chat-open': {
      const m = S.chat?.msgs?.find((x) => x.id === t.dataset.id);
      if (!m) return;
      S.view = 'mail';
      renderSidebar();
      renderListShell();
      return openMessage(m);
    }
    case 'chat-contact': {
      const p = S.chat.person;
      S.contact = (await findPeople(p.email).catch(() => []))[0] || { name: p.name, email: p.email, phones: [] };
      return openContacts();
    }
    case 'compose': return openCompose();
    case 'contacts': return openContacts();
    case 'settings': return openSettings();
    case 'refresh': return refreshList(false);
    case 'more': return loadMore(t);
    case 'open': {
      const m = S.list.find((x) => x.id === t.dataset.id);
      if (m) return openMessage(m);
      return;
    }
    case 'toggle-msg': {
      const art = t.closest('.msg');
      art.classList.toggle('open');
      if (art.classList.contains('open') && !art.dataset.mounted) {
        const m = S.thread.find((x) => x.id === art.dataset.mid);
        if (m) mountBody(m);
      }
      return;
    }
    case 'full-body':
    case 'short-body':
    case 'images': {
      const art = t.closest('.msg');
      const m = S.thread.find((x) => x.id === art.dataset.mid);
      if (!m) return;
      if (act === 'images') m.imagesAllowed = true;
      return mountBody(m, { full: act === 'full-body' || (act === 'images' && art.dataset.full === '1'), images: act === 'images' ? true : undefined });
    }
    case 'att': {
      t.disabled = true;
      try { downloadBlob(await S.api.attachmentBlob(t.dataset.mid, t.dataset.aid), t.dataset.name); } finally { t.disabled = false; }
      return;
    }
    case 'mark-unread': {
      const last = S.thread.filter((m) => !m.isDraft && m.from?.email !== myAddr()).at(-1) || S.sel;
      await S.api.markRead(last.id, false);
      last.isRead = false;
      const li = S.list.find((y) => y.id === last.id);
      if (li) li.isRead = false;
      bumpUnread(last.folderId, +1);
      renderItems();
      return toast('Marked as unread', 'ok', 1500);
    }
    case 'focus-reply': {
      $('#assistant')?.scrollIntoView({ behavior: 'smooth', block: 'end' });
      return setTimeout(() => ($('#draftText:not(:placeholder-shown)') || $('#revise') || $('#draftText'))?.focus(), 250);
    }
    case 'mode': {
      const st = replyState();
      st.mode = t.dataset.mode;
      return rerenderPanel();
    }
    case 'revise': return runDraft();
    case 'quick-revise': return runDraft(QUICK[t.dataset.rev]);
    case 'send-reply': return queueSend(false);
    case 'send-new': return queueSend(true);
    case 'undo-send': return undoSend();
    case 'outlook-reply': return replyInOutlook(false);
    case 'outlook-new': return replyInOutlook(true);
    case 'sched-menu': {
      const st = S.view === 'compose' ? S.compose : replyState();
      st.schedOpen = !st.schedOpen;
      return rerenderPanel();
    }
    case 'schedule': return queueSend(S.view === 'compose', new Date(t.dataset.at));
    case 'schedule-pick': {
      const v = $('#schedAt')?.value;
      if (!v) return $('#schedAt')?.focus();
      return queueSend(S.view === 'compose', new Date(v));
    }
    case 'signin-again': return auth.login();
    case 'reset-answers': {
      setCfg({ answers: '' });
      const ta = $('textarea[data-set="answers"]');
      ta.value = DEFAULT_ANSWERS;
      grow(ta);
      return flashSaved();
    }
    case 'pin': {
      const p = S.chat?.person.email === t.dataset.email ? S.chat.person : S.contact?.email === t.dataset.email ? S.contact : { name: t.dataset.name, email: t.dataset.email };
      togglePin(p);
      if (S.view === 'chats') { drawChat(); renderChatItems(); } else if (S.view === 'contacts') renderContact(S.contact);
      return;
    }
    case 'pick': {
      const input = $('#cTo');
      const parts = input.value.split(/[,;]/).slice(0, -1).map((s) => s.trim()).filter(Boolean);
      parts.push(`${t.dataset.name} <${t.dataset.email}>`);
      input.value = `${parts.join(', ')}, `;
      S.compose.to = input.value;
      $('#cSuggest').innerHTML = '';
      return input.focus();
    }
    case 'ctab':
      S.contactsTab = t.dataset.tab;
      $$('.tabs button').forEach((b) => b.classList.toggle('on', b === t));
      $('#citems').innerHTML = skeleton(6);
      return loadContacts();
    case 'contact': return renderContact(S.people[+t.dataset.i]);
    case 'write-to': return openCompose({ to: `${S.contact.name} <${S.contact.email}>, ` });
    case 'open-with': {
      const m = S.withList?.[+t.dataset.i];
      if (!m) return;
      S.view = 'mail';
      renderSidebar();
      renderListShell();
      return openMessage(m);
    }
    case 'toggle-key': {
      const k = $('#apiKey');
      k.type = k.type === 'password' ? 'text' : 'password';
      t.textContent = k.type === 'password' ? 'Show' : 'Hide';
      return;
    }
    case 'test-key': {
      const key = $('#apiKey').value.trim();
      if (!key) return toast('Paste your API key first.', 'err');
      setCfg({ apiKey: key });
      t.disabled = true;
      try {
        const models = await ai.listModels(key);
        toast(`Key works. ${models.length} models available.`, 'ok');
        fillModels();
      } catch (e) {
        toast(e.message, 'err', 6000);
      } finally {
        t.disabled = false;
      }
      return;
    }
    case 'reset-style': {
      setCfg({ style: '' });
      const ta = $('textarea[data-set="style"]');
      ta.value = ai.DEFAULT_STYLE;
      grow(ta);
      return flashSaved();
    }
    case 'export': {
      const { apiKey, ...rest } = cfg();
      return downloadBlob(new Blob([JSON.stringify(rest, null, 2)], { type: 'application/json' }), 'mail-assistant-settings.json');
    }
    default:
  }
}

function onInput(ev) {
  const el = ev.target;
  if (el.id === 'q') return runSearch(el.value);
  if (el.id === 'cq') { S.contactQ = el.value; return debounceContacts(); }
  if (el.id === 'chq') { S.chatQ = el.value; return renderChatItems(); }
  if (el.id === 'cTo') { S.compose.to = el.value; return suggest(el.value); }
  if (el.matches('textarea')) grow(el);
  if (el.id === 'draftText') { const st = S.view === 'compose' ? S.compose : replyState(); st.text = el.value; st.bracketOk = false; }
  if (el.id === 'sigBox') return saveField('signatureHtml', sanitizeBasic(el.innerHTML));
  if (el.dataset.set && el.type !== 'checkbox' && el.tagName !== 'SELECT') return saveField(el.dataset.set, el.value);
}
const debounceContacts = debounce(() => loadContacts(), 300);

async function onChange(ev) {
  const el = ev.target;
  if (el.id === 'mboxSel') return switchMailbox(el.value);
  if (el.id === 'answerSel') {
    const st = replyState();
    const a = [...answers(), GENERIC].find((x) => x.name === el.value);
    if (!a) return;
    Object.assign(st, { text: fillAnswer(a, replyTarget().from), source: 'standard', answer: a.name });
    return rerenderPanel();
  }
  if (el.type === 'radio' && el.dataset.set) {
    setCfg({ [el.dataset.set]: el.value });
    S.drafts = new Map();
    flashSaved();
    return keepScroll(() => openSettings());
  }
  if (el.dataset.set && (el.type === 'checkbox' || el.tagName === 'SELECT')) {
    setCfg({ [el.dataset.set]: el.type === 'checkbox' ? el.checked : el.value });
    if (el.dataset.set === 'useExamples') S.examples = null;
    flashSaved();
    if (/^acc|autoDraft/.test(el.dataset.set)) { S.drafts = new Map(); keepScroll(() => openSettings()); renderSidebar(); }
    return;
  }
  if (el.id === 'apiKey') { setCfg({ apiKey: el.value.trim() }); store.del('autoModel'); flashSaved(); return fillModels(); }
  if (el.id === 'importFile' && el.files?.[0]) {
    try {
      const data = JSON.parse(await el.files[0].text());
      delete data.apiKey;
      setCfg(Object.fromEntries(Object.entries(data).filter(([k]) => k in DEFAULTS)));
      toast('Settings imported', 'ok');
      openSettings();
    } catch {
      toast('That file is not a Mail Assistant settings export.', 'err');
    }
  }
}

function onKey(ev) {
  const el = ev.target;
  if (ev.key === 'Enter' && el.id === 'revise') {
    ev.preventDefault();
    if (el.value.trim()) runDraft(el.value.trim());
    return;
  }
  if (ev.key === 'Escape' && $('.sched-menu')) {
    const st = S.view === 'compose' ? S.compose : replyState();
    st.schedOpen = false;
    rerenderPanel();
  }
  if ((ev.key === 'Enter' || ev.key === ' ') && el.matches?.('.item[data-act]')) {
    ev.preventDefault();
    el.click();
  }
}

boot();

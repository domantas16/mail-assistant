import { CONFIG } from './config.js';
import {
  $, $$, esc, store, toast, debounce, avatar, fmtDate, fmtFull, fmtSize, textToHtml, htmlToText,
  cleanEmailHtml, srcdocFor, sanitizeBasic, copyRich, downloadBlob, grow, dayLabel, ICON,
} from './lib.js';
import * as auth from './auth.js';
import { graph } from './graph.js';
import { createDemo, cannedDraft, DEMO_SHARED } from './demo.js';
import * as ai from './ai.js';

// ---------- settings & state ----------
const DEFAULTS = {
  clientId: '', tenantId: '', apiKey: '', model: '', effort: 'low', myName: '', myRole: '',
  style: '', notes: '', useExamples: true, signatureHtml: '', addSignature: true,
  draftFont: 'Aptos, Calibri, Arial, sans-serif', loadImages: false, sharedMailboxes: '',
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
    chats: null, chat: null, chatQ: '', chatSeq: 0, mailbox: null,
  });
}
resetState();
const E = { app: $('#app'), side: $('#sidebar'), list: $('#listPane'), detail: $('#detailPane') };

// ---------- mailboxes: your own plus shared ones (info@, sales@ ...) ----------
const EMAIL = /[^\s<>,;"]+@[^\s<>,;"]+\.[^\s<>,;"]+/g;
function mailboxes() {
  if (S.mode === 'demo') return [DEMO_SHARED];
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
  const keep = { mode: S.mode, api: S.api, me: S.me, examples: S.examples };
  resetState();
  Object.assign(S, keep);
  S.mailbox = address || null;
  store.set('mailbox', S.mailbox);
  S.api.setMailbox(S.mailbox);
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
    <p class="lead">Read your Outlook mail in a clean view, and turn a short note into a polished reply in your own style.</p>
    ${err ? `<div class="alert err">${esc(friendlyAuthError(err))}</div>` : ''}
    <ol class="steps">
      <li class="step">
        <div class="step-h"><b>Connect Microsoft 365</b><span class="chip ${clientId ? 'ok' : 'wait'}">${clientId ? 'App IDs added' : 'Waiting for IT'}</span></div>
        <p class="muted">IT registers the app once and sends you two IDs. Paste them here:</p>
        <div class="grid2">
          <label>Application (client) ID<input id="wClient" value="${esc(clientId)}" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" spellcheck="false" autocomplete="off"></label>
          <label>Directory (tenant) ID<input id="wTenant" value="${esc(tenantId)}" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" spellcheck="false" autocomplete="off"></label>
        </div>
        <p class="muted small">Redirect URI for IT: <code>${esc(auth.redirectUri())}</code> <button class="link" data-act="copy-text" data-text="${esc(auth.redirectUri())}">Copy</button></p>
        <button class="btn primary" data-act="signin">${ICON.mail}<span>Sign in with Microsoft</span></button>
      </li>
      <li class="step">
        <div class="step-h"><b>Add your Claude API key</b><span class="chip ${c.apiKey ? 'ok' : 'wait'}">${c.apiKey ? 'Key saved' : 'Not added yet'}</span></div>
        <p class="muted">Create a key in the <a href="https://platform.claude.com/settings/keys" target="_blank" rel="noopener">Claude Console</a> (API keys). It is stored only in this browser and sent only to Anthropic.</p>
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

function renderListShell() {
  E.list.innerHTML = `
    <div class="list-head">
      <button class="icon-btn only-mobile" data-act="menu" aria-label="Menu">${ICON.menu}</button>
      <h2 id="listTitle">${esc(S.search ? 'Search results' : S.folder?.name || '')}</h2>
      <button class="icon-btn" data-act="refresh" title="Refresh" aria-label="Refresh">${ICON.refresh}</button>
    </div>
    <label class="search">${ICON.search}<input id="q" type="search" placeholder="Search all mail" value="${esc(S.search)}" autocomplete="off"></label>
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
  E.detail.innerHTML = empty(ICON.mail, 'Pick an email to read', 'Then write a few words about what you want to say and press <b>Draft reply</b>.');
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
        S.api.markRead(x.id, true).catch(() => {});
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
        <button class="btn subtle" data-act="focus-reply">${ICON.spark}<span>Reply with AI</span></button>
        ${S.sel.webLink ? `<a class="icon-btn" href="${esc(S.sel.webLink)}" target="_blank" rel="noopener" title="Open in Outlook" aria-label="Open in Outlook">${ICON.ext}</a>` : ''}
        <button class="icon-btn" data-act="mark-unread" title="Mark as unread" aria-label="Mark as unread">${ICON.unread}</button>
      </div>
    </div>
    <div class="msgs">${t.map((m, i) => msgHtml(m, open(m, i))).join('')}</div>
    ${replyPanel(target, st)}
  </div>`;
  t.forEach((m, i) => { if (open(m, i)) mountBody(m); });
  wireDraftArea();
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

// ---------- AI reply ----------
function replyTarget() {
  const real = S.thread.filter((m) => !m.isDraft);
  const others = real.filter((m) => m.from?.email && m.from.email !== myAddr());
  return others.at(-1) || real.at(-1) || S.sel;
}

function replyState() {
  const key = S.sel.conversationId || S.sel.id;
  if (!S.drafts.has(key)) {
    const t = replyTarget();
    const others = [...t.to, ...t.cc].filter((x) => x.email && x.email !== myAddr() && x.email !== t.from?.email);
    S.drafts.set(key, { note: '', text: '', history: null, mode: others.length ? 'all' : 'reply', busy: false, saved: null, meta: '' });
  }
  return S.drafts.get(key);
}

const QUICK = {
  Shorter: 'Make it shorter and tighter.',
  'More formal': 'Make it more formal.',
  Friendlier: 'Make it warmer and friendlier.',
  'Fix grammar only': 'Only fix grammar, spelling and flow. Keep the content and length.',
};

function draftTools(st, saveAct, sendAct) {
  return `
    <div class="draft ${st.text || st.busy ? '' : 'hidden'}">
      ${st.subjectField || ''}
      <textarea id="draftText" class="draft-text" spellcheck="true" aria-label="Draft">${esc(st.text)}</textarea>
      <div class="as-row wrap">
        <button class="btn" data-act="copy-draft">${ICON.copy}<span>Copy</span></button>
        <button class="btn" data-act="${saveAct}">${ICON.draft}<span>Save to drafts</span></button>
        <button class="btn ${st.confirmSend ? 'danger' : 'primary'}" data-act="${sendAct}">${ICON.send}<span>${st.confirmSend ? esc(st.confirmLabel) : 'Send'}</span></button>
        ${st.saved ? `<span class="saved">${S.mode === 'demo' ? 'Saved (demo)' : 'Saved to Drafts'}</span>${st.saved.webLink ? `<a class="btn subtle" href="${esc(st.saved.webLink)}" target="_blank" rel="noopener">${ICON.ext}<span>Open in Outlook</span></a>` : ''}` : ''}
        <span class="muted small meta">${esc(st.meta)}</span>
      </div>
      <div class="revise"><input id="revise" placeholder="Change something… e.g. “add that partial delivery is possible”" autocomplete="off"><button class="btn" data-act="revise">Revise</button></div>
      <div class="chips">${Object.keys(QUICK).map((k) => `<button class="chip-btn" data-act="quick-revise" data-rev="${esc(k)}">${esc(k)}</button>`).join('')}</div>
    </div>`;
}

function replyPanel(target, st) {
  const name = target.from?.name || target.from?.email || '';
  return `<section class="assistant ${st.busy ? 'busy' : ''}" id="assistant">
    <div class="as-head">
      <div class="as-title">${ICON.spark}<b>Reply with AI</b><span class="muted small">to ${esc(name)}${st.mode === 'all' ? ' and others' : ''}${S.mailbox ? `, from ${esc(S.mailbox)}` : ''}</span></div>
      <div class="seg" role="group" aria-label="Reply type">
        <button data-act="mode" data-mode="reply" class="${st.mode === 'reply' ? 'on' : ''}">Reply</button>
        <button data-act="mode" data-mode="all" class="${st.mode === 'all' ? 'on' : ''}">Reply all</button>
      </div>
    </div>
    <textarea id="note" rows="3" placeholder="What should the reply say? A few words in any language, e.g. “10 pcs 450 EUR each, 3 weeks, offer 5% for 20 pcs”">${esc(st.note)}</textarea>
    <div class="as-row">
      <button class="btn ai" data-act="draft" ${st.busy ? 'disabled' : ''}>${ICON.spark}<span>${st.busy ? 'Writing…' : st.text ? 'Draft again' : 'Draft reply'}</span></button>
      <span class="muted small">Ctrl + Enter</span>
    </div>
    ${draftTools(st, 'save-draft', 'send-reply')}
  </section>`;
}

function rerenderPanel() {
  if (S.view === 'compose') return renderCompose();
  const box = $('#assistant');
  if (!box || !S.sel) return;
  const keepNote = $('#note')?.value;
  box.outerHTML = replyPanel(replyTarget(), replyState());
  if (keepNote !== undefined && !replyState().busy) $('#note').value = keepNote;
  wireDraftArea();
}

function wireDraftArea() {
  ['#note', '#draftText', '#cNote'].forEach((s) => grow($(s)));
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

async function runDraft(revision) {
  const isNew = S.view === 'compose';
  const st = isNew ? S.compose : replyState();
  if (!st || st.busy) return;
  const c = cfg();
  const noteEl = $(isNew ? '#cNote' : '#note');
  if (noteEl) st.note = noteEl.value.trim();
  if (isNew) {
    st.to = $('#cTo')?.value || st.to;
    st.subject = $('#cSubject')?.value ?? st.subject;
    if (!st.note && !revision) return toast('Write a few words about what the email should say.', 'err');
  }
  const current = $('#draftText')?.value.trim() || st.text;
  const target = isNew ? null : replyTarget();

  if (!c.apiKey) {
    if (S.mode === 'demo') {
      const text = cannedDraft({ target, note: st.note, isNew });
      if (isNew) {
        const s = ai.splitSubject(text);
        st.subject = s.subject || st.subject;
        st.text = s.body;
      } else {
        st.text = text;
      }
      st.meta = 'Sample draft (no API key in demo)';
      st.saved = null;
      return rerenderPanel();
    }
    toast('Add your Claude API key in Settings first.', 'err');
    return openSettings('ai');
  }

  st.busy = true;
  rerenderPanel();
  const t0 = performance.now();
  try {
    const model = await chooseModel(c);
    const system = ai.buildSystem({ me: meForAI(), style: c.style || ai.DEFAULT_STYLE, notes: c.notes, examples: c.useExamples ? await examples() : [] });
    let messages;
    if (revision && st.history) {
      const prev = isNew ? `Subject: ${st.subject}\n\n${current}` : current;
      messages = [...st.history, { role: 'assistant', content: prev }, { role: 'user', content: ai.revisePrompt(revision, isNew) }];
    } else if (isNew) {
      messages = [{ role: 'user', content: ai.newEmailPrompt({ to: parseRecipients(st.to), note: st.note }) }];
    } else {
      const thread = S.thread.filter((m) => !m.isDraft).map((m) => ({ ...m, text: (m.text ??= htmlToText(m.bodyHtml)) }));
      messages = [{ role: 'user', content: ai.replyPrompt({ thread, target, note: st.note }) }];
    }
    const out = await ai.draft({
      key: c.apiKey, model, effort: c.effort, system, messages,
      onText: (t) => {
        const area = $('#draftText');
        if (area && (isNew ? S.view === 'compose' && S.compose === st : S.sel && replyState() === st)) {
          const first = !area.dataset.streaming;
          area.dataset.streaming = '1';
          area.closest('.draft')?.classList.remove('hidden');
          area.value = t;
          grow(area);
          if (first) area.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }
      },
    });
    if (isNew) {
      const { subject, body } = ai.splitSubject(out.text);
      st.subject = subject || st.subject;
      st.text = body;
    } else {
      st.text = out.text;
    }
    st.history = messages;
    st.saved = null;
    st.meta = `${ai.prettyModel(model)} · ${((performance.now() - t0) / 1000).toFixed(1)} s`;
  } catch (e) {
    fail(e);
  } finally {
    st.busy = false;
    const visible = isNew ? S.view === 'compose' && S.compose === st : (S.view === 'mail' || S.view === 'chats') && S.sel && replyState() === st;
    if (visible) rerenderPanel();
  }
}

function signatureHtml(c) {
  return c.addSignature && c.signatureHtml ? `<div>${sanitizeBasic(c.signatureHtml)}</div>` : '';
}

async function saveReplyDraft(btn) {
  const st = replyState();
  const text = $('#draftText')?.value.trim();
  if (!text) return toast('Nothing to save yet.');
  st.text = text;
  const c = cfg();
  btn.disabled = true;
  btn.lastChild.textContent = 'Saving…';
  try {
    const html = textToHtml(text, c.draftFont) + signatureHtml(c);
    st.saved = await S.api.createReplyDraft(replyTarget().id, { replyAll: st.mode === 'all', html });
    st.savedKey = `${st.mode}|${text}`;
    toast(S.mode === 'demo' ? 'Demo mode: saved to the sample Drafts folder.' : 'Saved to your Outlook Drafts. Review and send it from Outlook.', 'ok', 5000);
  } catch (e) {
    fail(e);
  }
  rerenderPanel();
}

async function copyDraft() {
  const text = $('#draftText')?.value.trim();
  if (!text) return;
  const ok = await copyRich(text, textToHtml(text, cfg().draftFont));
  toast(ok ? 'Copied. Paste it into your Outlook reply.' : 'Copy failed. Select the text and press Ctrl+C.', ok ? 'ok' : 'err');
}

// ---------- sending ----------
// First click arms the button ("Confirm: send to …"); a second click within 6 s sends.
async function sendDraft(btn, isNew) {
  const st = isNew ? S.compose : replyState();
  const text = $('#draftText')?.value.trim();
  if (!text) return toast('Nothing to send yet.');
  st.text = text;
  const c = cfg();
  let to = [];
  if (isNew) {
    st.to = $('#cTo').value;
    st.subject = $('#cSubject').value.trim();
    to = parseRecipients(st.to);
    if (!to.length) return toast('Add at least one recipient email address.', 'err');
    if (!st.subject) return toast('Add a subject first.', 'err');
  }
  if (!st.confirmSend) {
    const target = isNew ? null : replyTarget();
    const names = isNew ? to.map((x) => x.name.split(' ')[0]) : [target.from?.name || target.from?.email];
    const extra = isNew ? 0 : st.mode === 'all' ? [...target.to, ...target.cc].filter((x) => x.email && x.email !== myAddr() && x.email !== target.from?.email).length : 0;
    st.confirmSend = true;
    st.confirmLabel = `Confirm: send to ${names.slice(0, 2).join(', ')}${names.length > 2 || extra ? ` +${Math.max(names.length - 2, 0) + extra}` : ''}${c.addSignature && !c.signatureHtml ? ' (no signature set)' : ''}`;
    rerenderPanel();
    setTimeout(() => { if (st.confirmSend) { st.confirmSend = false; if ($('#assistant')) rerenderPanel(); } }, 6000);
    return;
  }
  st.confirmSend = false;
  btn.disabled = true;
  btn.lastChild.textContent = 'Sending…';
  const html = textToHtml(text, c.draftFont) + signatureHtml(c);
  try {
    const key = isNew ? `${st.to}|${st.subject}|${text}` : `${st.mode}|${text}`;
    let draft = st.saved && st.savedKey === key ? st.saved : null;
    if (!draft) {
      draft = isNew
        ? await S.api.createDraft({ to, subject: st.subject, html })
        : await S.api.createReplyDraft(replyTarget().id, { replyAll: st.mode === 'all', html });
    }
    await S.api.send(draft.id);
    toast(`Sent${S.mailbox ? ` from ${S.mailbox}` : ''}`, 'ok', 4000);
  } catch (e) {
    fail(e);
    return rerenderPanel();
  }
  if (isNew) {
    S.compose = { to: '', subject: '', note: '', text: '', busy: false, saved: null, meta: '', history: null };
    return renderCompose();
  }
  // Show the sent reply straight away (Outlook's search index can take a minute to catch up).
  const target = replyTarget();
  const sent = {
    id: `sent-${Date.now()}`, conversationId: target.conversationId, from: { name: S.mailbox || S.me?.name || 'Me', email: myAddr() },
    to: [target.from], cc: [], date: new Date().toISOString(), subject: `RE: ${target.subject.replace(/^re:\s*/i, '')}`,
    bodyHtml: html, preview: text.slice(0, 140), isRead: true, isDraft: false, hasAttachments: false,
  };
  Object.assign(st, { note: '', text: '', history: null, saved: null, savedKey: null, meta: '' });
  if (S.view === 'chats' && S.chat?.msgs) {
    S.chat.msgs.push(sent);
    const entry = S.chats?.find((x) => x.person.email === S.chat.person.email);
    if (entry) entry.last = { ...sent, out: true };
    drawChat();
    renderChatItems();
  } else {
    S.thread.push(sent);
    renderThread();
  }
}

// ---------- new email ----------
function openCompose(prefill = {}) {
  S.view = 'compose';
  S.compose = { to: prefill.to || '', subject: '', note: '', text: '', busy: false, saved: null, meta: '', history: null };
  renderSidebar();
  closeDrawer();
  E.app.classList.add('show-detail');
  if (!$('#items')) renderListShell();
  renderCompose();
  $(prefill.to ? '#cNote' : '#cTo')?.focus();
}

function renderCompose() {
  const c = S.compose;
  const st = { ...c, subjectField: `<label class="field">Subject<input id="cSubject" value="${esc(c.subject)}" autocomplete="off"></label>` };
  E.detail.innerHTML = `
  <div class="thread compose">
    <div class="thread-head"><button class="icon-btn only-mobile" data-act="back" aria-label="Back">${ICON.back}</button><h1>New email</h1></div>
    <section class="assistant ${c.busy ? 'busy' : ''}" id="assistant">
      <label class="field">To<input id="cTo" value="${esc(c.to)}" placeholder="Name or email. Start typing to search your contacts" autocomplete="off"></label>
      <div class="suggest" id="cSuggest"></div>
      <label class="field">What's it about?<textarea id="cNote" rows="3" placeholder="A few words in any language, e.g. “follow up on the 1030 nm mirror quote, 5% off for 20 pcs”">${esc(c.note)}</textarea></label>
      <div class="as-row"><button class="btn ai" data-act="draft" ${c.busy ? 'disabled' : ''}>${ICON.spark}<span>${c.busy ? 'Writing…' : c.text ? 'Draft again' : 'Draft email'}</span></button><span class="muted small">Ctrl + Enter</span></div>
      ${draftTools(st, 'save-new', 'send-new')}
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
    const people = (await S.api.people(token)).slice(0, 6);
    if (S.view !== 'compose') return;
    box.innerHTML = people.map((x) => `<button class="sugg" data-act="pick" data-name="${esc(x.name)}" data-email="${esc(x.email)}">${avatar(x, 24)}<span>${esc(x.name)}</span><small>${esc(x.email)}</small></button>`).join('');
  } catch { box.innerHTML = ''; }
}, 300);

async function saveNewDraft(btn) {
  const c = S.compose;
  c.to = $('#cTo').value;
  c.subject = $('#cSubject').value.trim();
  c.text = $('#draftText').value.trim();
  const to = parseRecipients(c.to);
  if (!to.length) return toast('Add at least one recipient email address.', 'err');
  if (!c.text) return toast('Nothing to save yet.');
  btn.disabled = true;
  btn.lastChild.textContent = 'Saving…';
  try {
    const conf = cfg();
    c.saved = await S.api.createDraft({ to, subject: c.subject, html: textToHtml(c.text, conf.draftFont) + signatureHtml(conf) });
    c.savedKey = `${c.to}|${c.subject}|${c.text}`;
    toast(S.mode === 'demo' ? 'Demo mode: saved to the sample Drafts folder.' : 'Saved to your Outlook Drafts.', 'ok', 5000);
  } catch (e) {
    fail(e);
  }
  renderCompose();
}

// ---------- chats: all emails with one person, messenger style ----------
async function openChats() {
  S.view = 'chats';
  renderSidebar();
  closeDrawer();
  E.list.innerHTML = `
    <div class="list-head">
      <button class="icon-btn only-mobile" data-act="menu" aria-label="Menu">${ICON.menu}</button>
      <h2>Chats</h2>
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
        <div class="prev">${c.last ? `${c.last.out ? '<span class="you">You: </span>' : ''}${esc(c.last.preview || c.last.subject)}` : `<span class="you">${esc(c.person.email)}</span>`}</div>
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
    S.api.markRead(m.id, true).catch(() => {});
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
        <div class="b-meta">${m.hasAttachments ? `<span class="b-att">${ICON.clip}</span>` : ''}${m.cc?.length ? `<span>+${m.cc.length} cc</span>` : ''}<button class="link" data-act="chat-open" data-id="${esc(m.id)}">Open email</button><time title="${esc(fmtFull(m.date))}">${esc(new Date(m.date).toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' }))}</time></div>
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
      <button data-act="ctab" data-tab="outlook" class="${S.contactsTab === 'outlook' ? 'on' : ''}">Saved contacts</button>
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
    if (S.contactsTab === 'outlook') {
      S.contacts ??= await S.api.contacts();
      list = q ? S.contacts.filter((x) => `${x.name} ${x.email} ${x.company}`.toLowerCase().includes(q)) : S.contacts;
    } else {
      list = await S.api.people(q);
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
      ${S.mode === 'live' ? `<p>Signed in as <b>${esc(S.me?.email || '')}</b>.</p>` : S.mode === 'demo' ? '<p class="muted">You are in demo mode. Add the IDs from IT, then use <b>Exit demo</b> in the sidebar and sign in.</p>' : ''}
      <div class="grid2">
        <label>Application (client) ID<input data-set="clientId" value="${esc(c.clientId)}" placeholder="${esc(CONFIG.clientId || 'from IT')}" spellcheck="false" autocomplete="off"></label>
        <label>Directory (tenant) ID<input data-set="tenantId" value="${esc(c.tenantId)}" placeholder="${esc(CONFIG.tenantId || 'from IT')}" spellcheck="false" autocomplete="off"></label>
      </div>
      <p class="muted small">Redirect URI for IT: <code>${esc(auth.redirectUri())}</code> <button class="link" data-act="copy-text" data-text="${esc(auth.redirectUri())}">Copy</button></p>
      <label>Shared mailboxes <span class="muted small">(one per line, e.g. info@company.com)</span>
        <textarea data-set="sharedMailboxes" rows="3" placeholder="info@company.com" spellcheck="false">${esc(c.sharedMailboxes)}</textarea></label>
      <p class="muted small">Use mailboxes you can already open in Outlook. Switch between them at the top of the sidebar. Sending from a shared mailbox also needs "Send as" rights on it.</p>
    </section>

    <section class="card" id="sec-ai"><h3>Claude (drafting)</h3>
      <label>API key
        <div class="row"><input id="apiKey" type="password" value="${esc(c.apiKey)}" placeholder="sk-ant-…" autocomplete="off" spellcheck="false"><button class="btn" data-act="toggle-key">Show</button><button class="btn" data-act="test-key">Test</button></div>
      </label>
      <p class="muted small">Stored only in this browser and sent only to Anthropic. Create one in the <a href="https://platform.claude.com/settings/keys" target="_blank" rel="noopener">Claude Console</a>. Each draft typically costs about one to three US cents.</p>
      <div class="grid2">
        <label>Model<select data-set="model" id="modelSel"><option value="">Automatic (newest Sonnet)</option>${c.model ? `<option value="${esc(c.model)}" selected>${esc(ai.prettyModel(c.model))}</option>` : ''}</select></label>
        <label>Speed<select data-set="effort">${[['low', 'Fast'], ['medium', 'Balanced'], ['high', 'Thorough']].map(([v, l]) => `<option value="${v}" ${c.effort === v ? 'selected' : ''}>${l}</option>`).join('')}</select></label>
      </div>
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

    <section class="card" id="sec-out"><h3>Outlook drafts</h3>
      <label class="check"><input type="checkbox" data-set="addSignature" ${c.addSignature ? 'checked' : ''}> Add my signature to drafts saved to Outlook</label>
      <p class="muted small">Outlook doesn't add your signature to drafts created by apps, so paste it below once (copy it from one of your sent emails). When you use <b>Copy</b> and paste into an Outlook reply instead, Outlook adds your signature as usual.</p>
      <div class="sig" id="sigBox" contenteditable="true" aria-label="Signature">${sanitizeBasic(c.signatureHtml || '')}</div>
      <div class="grid2">
        <label>Draft font<input data-set="draftFont" value="${esc(c.draftFont)}" autocomplete="off"></label>
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
  pendingSaves[key] = setTimeout(() => { setCfg({ [key]: value }); flashSaved(); if (key === 'sharedMailboxes') renderSidebar(); }, 350);
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
      S.contact = (await S.api.people(p.email).catch(() => []))[0] || { name: p.name, email: p.email, phones: [] };
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
      return setTimeout(() => $('#note')?.focus(), 250);
    }
    case 'mode': {
      const st = replyState();
      st.mode = t.dataset.mode;
      st.saved = null;
      return rerenderPanel();
    }
    case 'draft': return runDraft();
    case 'revise': {
      const v = $('#revise')?.value.trim();
      if (!v) return $('#revise')?.focus();
      return runDraft(v);
    }
    case 'quick-revise': return runDraft(QUICK[t.dataset.rev]);
    case 'copy-draft': return copyDraft();
    case 'save-draft': return saveReplyDraft(t);
    case 'save-new': return saveNewDraft(t);
    case 'send-reply': return sendDraft(t, false);
    case 'send-new': return sendDraft(t, true);
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
  if (el.id === 'note' && S.sel) replyState().note = el.value;
  if (el.id === 'cNote') S.compose.note = el.value;
  if (el.id === 'draftText') { const st = S.view === 'compose' ? S.compose : replyState(); st.text = el.value; }
  if (el.id === 'sigBox') return saveField('signatureHtml', sanitizeBasic(el.innerHTML));
  if (el.dataset.set && el.type !== 'checkbox' && el.tagName !== 'SELECT') return saveField(el.dataset.set, el.value);
}
const debounceContacts = debounce(() => loadContacts(), 300);

async function onChange(ev) {
  const el = ev.target;
  if (el.id === 'mboxSel') return switchMailbox(el.value);
  if (el.dataset.set && (el.type === 'checkbox' || el.tagName === 'SELECT')) {
    setCfg({ [el.dataset.set]: el.type === 'checkbox' ? el.checked : el.value });
    if (el.dataset.set === 'useExamples') S.examples = null;
    return flashSaved();
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
  if ((ev.ctrlKey || ev.metaKey) && ev.key === 'Enter' && (el.id === 'note' || el.id === 'cNote')) {
    ev.preventDefault();
    return runDraft();
  }
  if (ev.key === 'Enter' && el.id === 'revise') {
    ev.preventDefault();
    if (el.value.trim()) runDraft(el.value.trim());
    return;
  }
  if ((ev.key === 'Enter' || ev.key === ' ') && el.matches?.('.item[data-act]')) {
    ev.preventDefault();
    el.click();
  }
}

boot();

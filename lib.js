// Small shared helpers: DOM, storage, formatting, safe email rendering, icons.
export const $ = (sel, root = document) => root.querySelector(sel);
export const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

const ESC = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
export const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ESC[c]);

export const store = {
  get(key, fallback = null) {
    try { const v = localStorage.getItem('ma.' + key); return v == null ? fallback : JSON.parse(v); } catch { return fallback; }
  },
  set(key, value) { try { localStorage.setItem('ma.' + key, JSON.stringify(value)); } catch { /* storage full or blocked */ } },
  del(key) { try { localStorage.removeItem('ma.' + key); } catch { /* ignore */ } },
};

export function toast(msg, type = 'info', ms = 3500) {
  const box = $('#toasts');
  const t = document.createElement('div');
  t.className = `toast ${type}`;
  t.textContent = msg;
  box.append(t);
  setTimeout(() => { t.classList.add('out'); setTimeout(() => t.remove(), 300); }, ms);
}

export function debounce(fn, ms = 300) {
  let t;
  return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); };
}

const PALETTE = ['#3156d3', '#0f8a6a', '#b4531a', '#7a46c6', '#c23a5b', '#1f7fb3', '#5d6b1f', '#9b4d1f', '#2d6a8f', '#8a3fa0'];
export function colorFor(s = '') {
  let h = 0;
  for (const c of s) h = (h * 31 + c.charCodeAt(0)) >>> 0;
  return PALETTE[h % PALETTE.length];
}
export function initials(name = '', email = '') {
  const src = (name || email.split('@')[0] || '?').replace(/\b(dr|prof|mr|mrs|ms)\.?\s/gi, ' ').replace(/[^\p{L}\p{N} ]/gu, ' ').trim();
  const parts = src.split(/\s+/).filter(Boolean);
  return ((parts[0]?.[0] || '?') + (parts.length > 1 ? parts[parts.length - 1][0] : '')).toUpperCase();
}
export function avatar(p, size = 36) {
  const key = p?.email || p?.name || '';
  return `<span class="av" style="--c:${colorFor(key)};--s:${size}px" aria-hidden="true">${esc(initials(p?.name, p?.email))}</span>`;
}

const LOC = 'en-GB';
export function fmtDate(iso) {
  if (!iso) return '';
  const d = new Date(iso), now = new Date();
  const time = d.toLocaleTimeString(LOC, { hour: '2-digit', minute: '2-digit' });
  if (d.toDateString() === now.toDateString()) return time;
  if ((now - d) / 864e5 < 6) return `${d.toLocaleDateString(LOC, { weekday: 'short' })} ${time}`;
  if (d.getFullYear() === now.getFullYear()) return d.toLocaleDateString(LOC, { day: 'numeric', month: 'short' });
  return d.toLocaleDateString(LOC, { day: 'numeric', month: 'short', year: 'numeric' });
}
export const fmtFull = (iso) => (iso ? new Date(iso).toLocaleString(LOC, { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '');
export function fmtSize(n = 0) {
  if (n < 1024) return `${n} B`;
  if (n < 1048576) return `${Math.round(n / 1024)} KB`;
  return `${(n / 1048576).toFixed(1)} MB`;
}

// Plain-text draft -> Outlook-friendly HTML (paragraphs, line breaks, clickable links).
const URL_RE = /\b(?:https?:\/\/|www\.)[^\s<>"']+[^\s<>"'.,;:!?)\]]/g;
function linkify(line) {
  let out = '', last = 0;
  for (const m of line.matchAll(URL_RE)) {
    const href = m[0].startsWith('www.') ? `https://${m[0]}` : m[0];
    out += esc(line.slice(last, m.index)) + `<a href="${esc(href)}">${esc(m[0])}</a>`;
    last = m.index + m[0].length;
  }
  return out + esc(line.slice(last));
}
export function textToHtml(text, font = 'Aptos, Calibri, Arial, sans-serif') {
  const paras = String(text).replace(/\r\n/g, '\n').trim().split(/\n{2,}/);
  const body = paras.map((p) => `<p style="margin:0 0 11pt 0">${p.split('\n').map(linkify).join('<br>')}</p>`).join('');
  return `<div style="font-family:${esc(font)};font-size:11pt;color:#000000">${body}</div>`;
}
export const textBodyToHtml = (t) => `<div style="white-space:pre-wrap">${linkify(String(t || ''))}</div>`;

// Email HTML -> readable plain text (used for the AI prompt).
const BLOCK = /^(P|DIV|LI|TR|H[1-6]|BLOCKQUOTE|PRE|TABLE|UL|OL|HR|SECTION|ARTICLE|HEADER|FOOTER)$/;
export function htmlToText(html) {
  if (!html) return '';
  const doc = new DOMParser().parseFromString(html, 'text/html');
  doc.querySelectorAll('style,script,head,title,img').forEach((n) => n.remove());
  let out = '';
  const walk = (n) => {
    if (n.nodeType === 3) { out += n.nodeValue.replace(/\s+/g, ' '); return; }
    if (n.nodeType !== 1) return;
    const tag = n.tagName;
    if (tag === 'BR') { out += '\n'; return; }
    if (tag === 'LI') out += '\n- ';
    else if (BLOCK.test(tag)) out += '\n';
    if (tag === 'TD' || tag === 'TH') out += ' ';
    n.childNodes.forEach(walk);
    if (tag === 'A') {
      const href = n.getAttribute('href') || '';
      if (/^https?:/i.test(href) && !n.textContent.includes(href.replace(/^https?:\/\//, '').slice(0, 20))) out += ` (${href})`;
    }
    if (BLOCK.test(tag)) out += '\n';
  };
  walk(doc.body);
  return out.replace(/[ \t\u00a0]+\n/g, '\n').replace(/\n[ \t\u00a0]+/g, '\n').replace(/\n{3,}/g, '\n\n').replace(/[ \t\u00a0]{2,}/g, ' ').trim();
}

// Remove active content from email HTML and swap inline (cid:) images for data URLs.
// The result is shown inside a sandboxed iframe with no scripts, so this is a second layer.
export function cleanEmailHtml(html, cids = {}) {
  const doc = new DOMParser().parseFromString(html || '', 'text/html');
  doc.querySelectorAll('script,meta,base,iframe,frame,frameset,object,embed,form,link,noscript').forEach((n) => n.remove());
  let remoteImages = false;
  doc.querySelectorAll('*').forEach((el) => {
    for (const a of [...el.attributes]) {
      if (/^on/i.test(a.name) || /^\s*javascript:/i.test(a.value)) el.removeAttribute(a.name);
    }
  });
  doc.querySelectorAll('img').forEach((img) => {
    const src = img.getAttribute('src') || '';
    if (/^cid:/i.test(src)) {
      const data = cids[src.slice(4)] || cids[decodeURIComponent(src.slice(4))];
      if (data) img.setAttribute('src', data);
    } else if (/^https?:/i.test(src)) remoteImages = true;
  });
  if (/url\(\s*['"]?https?:/i.test(html || '')) remoteImages = true;
  const styles = [...doc.head.querySelectorAll('style')].map((s) => s.outerHTML).join('');
  return { html: styles + doc.body.innerHTML, remoteImages };
}

export function srcdocFor(html, allowImages) {
  const img = allowImages ? 'https: data: blob:' : 'data: blob:';
  return `<!doctype html><html><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; img-src ${img}; font-src data:">
<base target="_blank">
<style>html,body{margin:0;background:#fff;color:#1d2433}body{padding:2px 2px 6px;font:14px/1.55 "Segoe UI",system-ui,-apple-system,Roboto,sans-serif;overflow-wrap:anywhere}
img{max-width:100%;height:auto}pre{white-space:pre-wrap}a{color:#2f55d4}
blockquote{margin:0 0 0 .4em;padding-left:.8em;border-left:3px solid #d0d5dd;color:#475467}</style>
</head><body>${html}</body></html>`;
}

// Strip active content from user-pasted HTML (e.g. the Outlook signature).
export function sanitizeBasic(html) {
  return cleanEmailHtml(html).html;
}

export async function copyRich(text, html) {
  try {
    if (html && window.ClipboardItem) {
      await navigator.clipboard.write([new ClipboardItem({
        'text/plain': new Blob([text], { type: 'text/plain' }),
        'text/html': new Blob([html], { type: 'text/html' }),
      })]);
    } else {
      await navigator.clipboard.writeText(text);
    }
    return true;
  } catch {
    try { await navigator.clipboard.writeText(text); return true; } catch { return false; }
  }
}

export function downloadBlob(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

// Group messages into one "chat" per person (the other side of each email).
export function groupChats(msgs, me) {
  const map = new Map();
  for (const m of msgs) {
    if (m.isDraft) continue;
    const out = m.from?.email === me;
    const p = out ? m.to[0] : m.from;
    if (!p?.email || p.email === me) continue;
    const g = map.get(p.email) || { person: { name: p.name || p.email, email: p.email }, last: null, unread: 0, count: 0 };
    g.count++;
    if (!out && !m.isRead) g.unread++;
    if (!g.last || new Date(m.date) > new Date(g.last.date)) g.last = { ...m, out };
    if (!out && p.name && p.name.toLowerCase() !== p.email) g.person.name = p.name;
    map.set(p.email, g);
  }
  return [...map.values()].sort((a, b) => new Date(b.last.date) - new Date(a.last.date));
}

export function dayLabel(iso) {
  const d = new Date(iso), now = new Date();
  const y = new Date(now); y.setDate(now.getDate() - 1);
  if (d.toDateString() === now.toDateString()) return 'Today';
  if (d.toDateString() === y.toDateString()) return 'Yesterday';
  return d.toLocaleDateString(LOC, { weekday: 'short', day: 'numeric', month: 'short', ...(d.getFullYear() !== now.getFullYear() ? { year: 'numeric' } : {}) });
}

export function grow(ta) {
  if (!ta) return;
  ta.style.height = 'auto';
  ta.style.height = `${Math.min(ta.scrollHeight + 2, 640)}px`;
}

const svg = (d) => `<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${d}</svg>`;
export const ICON = {
  logo: '<svg viewBox="0 0 32 32" width="28" height="28" aria-hidden="true"><defs><linearGradient id="ma-lg" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#4f6ef7"/><stop offset="1" stop-color="#8a4bd6"/></linearGradient></defs><rect width="32" height="32" rx="9" fill="url(#ma-lg)"/><path d="M7.5 11h17v11h-17z" fill="none" stroke="#fff" stroke-width="1.8" stroke-linejoin="round"/><path d="M8 11.6l8 5.9 8-5.9" fill="none" stroke="#fff" stroke-width="1.8" stroke-linejoin="round"/><path d="M25 3.8l.8 1.9 1.9.8-1.9.8-.8 1.9-.8-1.9-1.9-.8 1.9-.8z" fill="#fff"/></svg>',
  inbox: svg('<path d="M22 12h-6l-2 3h-4l-2-3H2"/><path d="M5.5 5h13L22 12v6a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2v-6z"/>'),
  send: svg('<path d="M22 2 11 13"/><path d="M22 2 15 22l-4-9-9-4z"/>'),
  draft: svg('<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="M8 13h8M8 17h5"/>'),
  archive: svg('<rect x="2" y="3" width="20" height="5" rx="1"/><path d="M4 8v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8"/><path d="M10 12h4"/>'),
  folder: svg('<path d="M3 6a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>'),
  trash: svg('<path d="M3 6h18M8 6V4h8v2M6 6l1 14h10l1-14"/>'),
  chat: svg('<path d="M21 12a8.5 8.5 0 0 1-12.4 7.6L3 21l1.4-5.4A8.5 8.5 0 1 1 21 12z"/><path d="M8.5 10.5h7M8.5 14h4.5"/>'),
  people: svg('<circle cx="9" cy="8" r="4"/><path d="M2 21a7 7 0 0 1 14 0"/><path d="M16 4a4 4 0 0 1 0 8M22 21a7 7 0 0 0-4-6.3"/>'),
  gear: svg('<circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M4.9 4.9l2.1 2.1M17 17l2.1 2.1M2 12h3M19 12h3M4.9 19.1 7 17M17 7l2.1-2.1"/>'),
  spark: svg('<path d="M12 3l1.9 5.1L19 10l-5.1 1.9L12 17l-1.9-5.1L5 10l5.1-1.9z"/><path d="M19 15l.8 2.2L22 18l-2.2.8L19 21l-.8-2.2L16 18l2.2-.8z"/>'),
  pen: svg('<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z"/>'),
  search: svg('<circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/>'),
  refresh: svg('<path d="M21 12a9 9 0 1 1-2.6-6.4L21 8"/><path d="M21 3v5h-5"/>'),
  copy: svg('<rect x="9" y="9" width="12" height="12" rx="2"/><path d="M5 15H4a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1v1"/>'),
  ext: svg('<path d="M14 4h6v6"/><path d="M10 14 20 4"/><path d="M19 13v6a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V6a1 1 0 0 1 1-1h6"/>'),
  back: svg('<path d="m15 18-6-6 6-6"/>'),
  menu: svg('<path d="M4 6h16M4 12h16M4 18h16"/>'),
  clip: svg('<path d="m21.4 11.1-9.2 9.2a6 6 0 0 1-8.5-8.5l9.2-9.2a4 4 0 0 1 5.7 5.7l-9.2 9.2a2 2 0 0 1-2.8-2.8l8.5-8.5"/>'),
  flag: svg('<path d="M4 22V4"/><path d="M4 4h12l-2 4 2 4H4"/>'),
  mail: svg('<rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-10 6L2 7"/>'),
  phone: svg('<path d="M5 3h4l2 5-2.5 1.5a11 11 0 0 0 6 6L16 13l5 2v4a2 2 0 0 1-2 2A17 17 0 0 1 3 5a2 2 0 0 1 2-2"/>'),
  building: svg('<rect x="4" y="2" width="16" height="20" rx="2"/><path d="M9 22v-4h6v4M8 6h2M14 6h2M8 10h2M14 10h2M8 14h2M14 14h2"/>'),
  eye: svg('<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/>'),
  pin: svg('<path d="M12 17v5"/><path d="M8.5 3h7l-1 6.5 3 2.5v2h-11v-2l3-2.5z"/>'),
  unread: svg('<path d="M14 4H4a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-7"/><path d="m2 7 10 6 4.5-2.7"/><circle cx="19.5" cy="5.5" r="2.5" fill="currentColor"/>'),
};

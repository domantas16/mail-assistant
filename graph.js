// Outlook access through Microsoft Graph, as the signed-in user.
import { getAccessToken } from './auth.js?v=20261119b';
import { htmlToText, textBodyToHtml, groupChats } from './lib.js?v=20261119b';

const BASE = 'https://graph.microsoft.com/v1.0';
const LIST = 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,bodyPreview,isRead,isDraft,hasAttachments,conversationId,flag,importance,webLink,parentFolderId';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
let known = {}; // well-known folder ids of the current mailbox: inbox, drafts, sentitems, ...
let MB = '/me'; // current mailbox: your own, or a shared one via /users/{address}

async function call(path, { method = 'GET', body, raw = false } = {}) {
  const url = path.startsWith('https://') ? path : BASE + path;
  for (let attempt = 0; ; attempt++) {
    const token = await getAccessToken();
    const res = await fetch(url, {
      method,
      headers: { Authorization: `Bearer ${token}`, ...(body !== undefined ? { 'Content-Type': 'application/json' } : {}) },
      body: body !== undefined ? JSON.stringify(body) : undefined,
    });
    if ([429, 503, 504].includes(res.status) && attempt < 3) { await sleep(1500 * (attempt + 1)); continue; }
    if (!res.ok) {
      let code = '', msg = `${res.status} ${res.statusText}`;
      try { const j = await res.json(); code = j.error?.code || ''; msg = j.error?.message || msg; } catch { /* no body */ }
      const e = new Error(res.status === 403
        ? `Microsoft 365 refused this request (${code || 'Forbidden'}). A permission may still need IT approval.`
        : `Outlook: ${msg}`);
      e.status = res.status;
      e.code = res.status === 401 ? 'auth_required' : code;
      throw e;
    }
    if (raw) return res;
    const txt = await res.text(); // send returns 202 with an empty body
    return txt ? JSON.parse(txt) : null;
  }
}

const addr = (r) => (r?.emailAddress ? {
  name: r.emailAddress.name || r.emailAddress.address || '',
  email: (r.emailAddress.address || '').toLowerCase(),
} : null);

function mapMsg(m) {
  return {
    id: m.id,
    subject: m.subject || '(no subject)',
    from: addr(m.from) || { name: '(unknown sender)', email: '' },
    to: (m.toRecipients || []).map(addr).filter(Boolean),
    cc: (m.ccRecipients || []).map(addr).filter(Boolean),
    date: m.receivedDateTime || m.sentDateTime,
    preview: (m.bodyPreview || '').trim(),
    isRead: m.isRead !== false,
    isDraft: !!m.isDraft,
    hasAttachments: !!m.hasAttachments,
    conversationId: m.conversationId,
    flagged: m.flag?.flagStatus === 'flagged',
    important: m.importance === 'high',
    webLink: m.webLink,
    folderId: m.parentFolderId,
  };
}

const bodyHtml = (b) => (b?.contentType === 'text' ? textBodyToHtml(b.content) : b?.content || '');
const toThreadMsg = (m) => ({ ...mapMsg(m), bodyHtml: bodyHtml(m.uniqueBody) });
const quote = (s) => encodeURIComponent(`"${String(s).replace(/"/g, '')}"`);
const SIGNOFF = /^\s*(best|kind|warm|many thanks|thanks and)\s+(regards|wishes)|^\s*(regards|cheers|thank you|thanks|br),?\s*$|^\s*pagarbiai|^\s*su pagarba/i;

export const graph = {
  // Switch between your own mailbox (null) and a shared one, e.g. 'info@company.com'.
  setMailbox(address) {
    MB = address ? `/users/${encodeURIComponent(address)}` : '/me';
    known = {};
  },

  // Sends a draft created by createReplyDraft or createDraft. With `at` (a Date), Exchange keeps it
  // in the Outbox and sends it at that time (deferred delivery, the same as Outlook's "Schedule send").
  async send(draftId, { at } = {}) {
    if (at) {
      await call(`${MB}/messages/${draftId}`, {
        method: 'PATCH',
        body: { singleValueExtendedProperties: [{ id: 'SystemTime 0x3FEF', value: new Date(at).toISOString() }] },
      });
    }
    await call(`${MB}/messages/${draftId}/send`, { method: 'POST' });
  },

  async me() {
    const u = await call('/me?$select=displayName,mail,userPrincipalName,jobTitle');
    return { name: u.displayName || '', email: (u.mail || u.userPrincipalName || '').toLowerCase(), jobTitle: u.jobTitle || '' };
  },

  async folders() {
    const wk = ['inbox', 'drafts', 'sentitems', 'archive', 'junkemail', 'deleteditems'];
    const got = await Promise.all(wk.map((w) => call(`${MB}/mailFolders/${w}?$select=id,displayName,unreadItemCount`).catch(() => null)));
    wk.forEach((w, i) => { if (got[i]) known[w] = got[i].id; });
    const labels = { inbox: ['Inbox', 'inbox'], drafts: ['Drafts', 'draft'], sentitems: ['Sent', 'send'], archive: ['Archive', 'archive'], junkemail: ['Junk', 'folder'], deleteditems: ['Deleted', 'trash'] };
    const make = (w, i) => ({ id: got[i].id, key: w, name: labels[w][0], icon: labels[w][1], unread: w === 'inbox' || w === 'junkemail' ? got[i].unreadItemCount : 0 });
    const main = wk.slice(0, 4).map((w, i) => (got[i] ? make(w, i) : null)).filter(Boolean);
    const tail = wk.slice(4).map((w, j) => (got[j + 4] ? make(w, j + 4) : null)).filter(Boolean);
    const skip = new Set(Object.values(known));
    const HIDE = /^(outbox|conversation history|sync issues.*|rss( feeds| subscriptions)?|scheduled|social activity notifications|yammer root|files|externalcontacts|personmetadata|quick step settings|clutter|notes|journal|tasks|calendar|contacts)$/i;
    const [top, kids] = await Promise.all([
      call(`${MB}/mailFolders?$top=100&$select=id,displayName,unreadItemCount`).catch(() => ({ value: [] })),
      known.inbox ? call(`${MB}/mailFolders/inbox/childFolders?$top=100&$select=id,displayName,unreadItemCount`).catch(() => ({ value: [] })) : { value: [] },
    ]);
    const fold = (f, child) => ({ id: f.id, name: f.displayName, icon: 'folder', unread: f.unreadItemCount, child });
    const other = [
      ...(kids.value || []).map((f) => fold(f, true)),
      ...(top.value || []).filter((f) => !skip.has(f.id) && !HIDE.test(f.displayName)).map((f) => fold(f, false)),
      ...tail,
    ];
    return { main, other };
  },

  async messages(folderId, { search, next } = {}) {
    let url = next;
    if (!url) {
      url = search
        ? `${MB}/messages?$search=${quote(search)}&$top=30&$select=${LIST}`
        : `${MB}/mailFolders/${encodeURIComponent(folderId || 'inbox')}/messages?$top=40&$orderby=receivedDateTime desc&$select=${LIST}`;
    }
    const r = await call(url);
    return { items: (r.value || []).map(mapMsg), next: r['@odata.nextLink'] || null };
  },

  async thread(msg) {
    let items = [];
    if (msg.conversationId) {
      const filter = encodeURIComponent(`conversationId eq '${msg.conversationId.replace(/'/g, "''")}'`);
      const r = await call(`${MB}/messages?$filter=${filter}&$top=50&$select=${LIST},uniqueBody`);
      items = (r.value || []).filter((m) => m.parentFolderId !== known.deleteditems).map(toThreadMsg);
    }
    if (!items.some((m) => m.id === msg.id)) {
      items.push(toThreadMsg(await call(`${MB}/messages/${msg.id}?$select=${LIST},uniqueBody`)));
    }
    return items.sort((a, b) => new Date(a.date) - new Date(b.date));
  },

  // One chat per person, built from recent received and sent mail.
  async chats(myEmail) {
    const [inbox, sent] = await Promise.all([
      call(`${MB}/mailFolders/inbox/messages?$top=60&$orderby=receivedDateTime desc&$select=${LIST}`),
      call(`${MB}/mailFolders/sentitems/messages?$top=60&$orderby=sentDateTime desc&$select=${LIST}`).catch(() => ({ value: [] })),
    ]);
    return groupChats([...(inbox.value || []), ...(sent.value || [])].map(mapMsg), myEmail);
  },

  // Every email with one person, oldest first, with readable bodies.
  async chat(email) {
    const r = await call(`${MB}/messages?$search=${quote(`participants:${email}`)}&$top=40&$select=${LIST},uniqueBody`);
    return (r.value || [])
      .filter((m) => m.parentFolderId !== known.deleteditems && !m.isDraft)
      .map(toThreadMsg)
      .sort((a, b) => new Date(a.date) - new Date(b.date));
  },

  async fullBody(id) {
    return bodyHtml((await call(`${MB}/messages/${id}?$select=body`)).body);
  },

  async attachments(id) {
    const r = await call(`${MB}/messages/${id}/attachments?$select=id,name,contentType,size,isInline`);
    return (r.value || []).map((a) => ({ id: a.id, name: a.name || 'attachment', type: a.contentType || '', size: a.size || 0, inline: !!a.isInline }));
  },

  async inlineImages(id) {
    const list = await graph.attachments(id);
    const imgs = list.filter((a) => /^image\//i.test(a.type) && a.size < 3e6).slice(0, 15);
    const map = {};
    await Promise.all(imgs.map(async (a) => {
      try {
        const f = await call(`${MB}/messages/${id}/attachments/${a.id}`);
        if (f.contentId && f.contentBytes) map[f.contentId.replace(/^<|>$/g, '')] = `data:${f.contentType || a.type};base64,${f.contentBytes}`;
      } catch { /* skip this image */ }
    }));
    return map;
  },

  async attachmentBlob(msgId, attId) {
    return (await call(`${MB}/messages/${msgId}/attachments/${attId}/$value`, { raw: true })).blob();
  },

  async markRead(id, isRead = true) {
    await call(`${MB}/messages/${id}`, { method: 'PATCH', body: { isRead } });
  },

  // Creates a reply draft in Outlook (with the quoted original below) and puts our text on top.
  async createReplyDraft(messageId, { replyAll, html }) {
    const draft = await call(`${MB}/messages/${messageId}/${replyAll ? 'createReplyAll' : 'createReply'}`, { method: 'POST', body: {} });
    const orig = draft.body?.contentType === 'text' ? textBodyToHtml(draft.body.content) : draft.body?.content || '';
    const content = /<body[^>]*>/i.test(orig) ? orig.replace(/<body[^>]*>/i, (m) => `${m}${html}`) : html + orig;
    const updated = await call(`${MB}/messages/${draft.id}`, { method: 'PATCH', body: { body: { contentType: 'HTML', content } } });
    return { id: draft.id, webLink: updated?.webLink || draft.webLink };
  },

  async createDraft({ to, subject, html }) {
    const m = await call(`${MB}/messages`, {
      method: 'POST',
      body: { subject, body: { contentType: 'HTML', content: html }, toRecipients: to.map((p) => ({ emailAddress: { address: p.email, name: p.name || p.email } })) },
    });
    return { id: m.id, webLink: m.webLink };
  },

  async people(q = '') {
    const sel = '$select=displayName,scoredEmailAddresses,companyName,jobTitle,phones,personType';
    const r = await call(q ? `/me/people?$search=${quote(q)}&$top=25&${sel}` : `/me/people?$top=60&${sel}`);
    return (r.value || [])
      .filter((p) => p.scoredEmailAddresses?.[0]?.address && (!p.personType || p.personType.class === 'Person'))
      .map((p) => ({
        name: p.displayName || p.scoredEmailAddresses[0].address,
        email: p.scoredEmailAddresses[0].address.toLowerCase(),
        company: p.companyName || '',
        jobTitle: p.jobTitle || '',
        phones: (p.phones || []).map((x) => x.number).filter(Boolean),
      }));
  },

  async contacts() {
    let url = '/me/contacts?$top=200&$orderby=displayName&$select=displayName,emailAddresses,companyName,jobTitle,businessPhones,mobilePhone';
    const out = [];
    for (let i = 0; url && i < 5; i++) {
      const r = await call(url);
      out.push(...(r.value || []));
      url = r['@odata.nextLink'];
    }
    return out.map((c) => ({
      name: c.displayName || c.emailAddresses?.[0]?.address || '(no name)',
      email: (c.emailAddresses?.[0]?.address || '').toLowerCase(),
      emails: (c.emailAddresses || []).map((e) => e.address).filter(Boolean),
      company: c.companyName || '',
      jobTitle: c.jobTitle || '',
      phones: [...(c.businessPhones || []), c.mobilePhone].filter(Boolean),
    }));
  },

  // A few recent replies to people outside the company, signature removed: used as tone examples.
  async sentExamples(n = 4, myEmail = '') {
    const domain = myEmail.split('@')[1] || '';
    const r = await call(`/me/mailFolders/sentitems/messages?$top=30&$orderby=sentDateTime desc&$select=subject,toRecipients,uniqueBody`);
    const out = [];
    for (const m of r.value || []) {
      const to = (m.toRecipients || []).map(addr).filter(Boolean);
      if (!to.length || (domain && to.every((t) => t.email.endsWith(`@${domain}`)))) continue;
      const lines = (m.uniqueBody?.contentType === 'text' ? m.uniqueBody.content : htmlToText(m.uniqueBody?.content || '')).trim().split('\n');
      const cut = lines.findIndex((l, i) => i > 1 && SIGNOFF.test(l));
      const text = (cut > 0 ? lines.slice(0, cut) : lines).join('\n').trim();
      if (text.length < 120 || text.length > 2500) continue;
      out.push({ to: to.map((t) => t.name).join(', '), subject: m.subject || '', text });
      if (out.length >= n) break;
    }
    return out;
  },
};

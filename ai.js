// Drafting with Claude through the Anthropic API, called directly from the browser
// with the user's own API key (kept in this browser only).
const API = 'https://api.anthropic.com/v1';
export const FALLBACK_MODEL = 'claude-sonnet-5-5';

const headers = (key) => ({
  'x-api-key': key,
  'anthropic-version': '2023-06-01',
  'content-type': 'application/json',
  'anthropic-dangerous-direct-browser-access': 'true',
});

export const DEFAULT_STYLE = `Voice: warm, professional and concise. Formal but not stiff, plain clear English, no heavy jargon. Make the point, then stop: usually 3-8 short sentences in short paragraphs.

Language: write in the language of the email being answered (usually English), even when my note is in Lithuanian.

Greeting: "Dear [first name]," or "Hi [first name]," if they already write informally. For customers, a warm opener fits, e.g. "Great to hear from you!" or "Thank you for reaching out!"

Customer inquiries: reassure where natural ("You came to the right place"), give clear, structured guidance with concrete numbers (quantities, thresholds, options), and add a direct link with one line of context when it helps. A light emoji (😊) is fine now and then in friendly customer emails, never in formal ones or complaints.

OPTOSHOP: when pointing customers to the shop, say "You can order, or request for a quotation directly in OPTOSHOP." (https://optoshop.optoman.com). Standard shipping on OPTOSHOP quotes is $75 USD.
Custom / technical requests: point to OPTOguide (https://www.optoman.com/optoguide/) or ask for the specs: wavelength, AOI, polarization, fluence or power, pulse width, repetition rate, beam size, quantity.
Classification when asked: HTS 9001.90.9000, ECCN EAR99.

Closing: end with a short friendly line when it fits, e.g. "Keep in touch!" or "Let me know if you have any questions."`;

const RULES = `Rules:
- Output only the email text: no commentary, no markdown, no quotation marks around it.
- No signature block, no sign-off such as "Best regards" or "Kind regards", and no name at the end. The signature is added separately.
- Use only facts from my note, the thread, or the notes above. Never invent prices, lead times, specifications, dates or commitments. If something essential is missing, put a placeholder in square brackets, e.g. [lead time].
- Never promise to send curves, data, samples, documents or a call unless my note says so.
- Keep names, numbers, part numbers and links exactly as given.`;

export function buildSystem({ me, style, notes, examples }) {
  const who = `${me.name || 'me'}${me.role ? `, ${me.role}` : ''} at OPTOMAN (precision laser optics, Vilnius, Lithuania; online shop OPTOSHOP)`;
  const parts = [
    `You write emails for ${who}. I give you a short note (often in Lithuanian, sometimes just keywords) saying what the email should say; you turn it into the finished email I will send.`,
    `<style_guide>\n${style}\n</style_guide>`,
  ];
  if (notes?.trim()) parts.push(`<my_notes_and_templates>\n${notes.trim()}\n</my_notes_and_templates>\nUse these only when relevant to the email at hand.`);
  if (examples?.length) {
    parts.push(`<my_recent_emails>\n${examples.map((x) => `To: ${x.to}\nSubject: ${x.subject}\n\n${x.text}`).join('\n\n---\n\n')}\n</my_recent_emails>\nMatch the voice of these examples; never copy their content.`);
  }
  parts.push(RULES);
  return parts.join('\n\n');
}

const who = (p) => (p ? (p.name && p.name.toLowerCase() !== p.email ? `${p.name} <${p.email}>` : p.email) : '');

export function replyPrompt({ thread, target, note, standard }) {
  const blocks = [];
  let total = 0;
  for (const m of [...thread].reverse()) {
    const block = [
      `From: ${who(m.from)}`,
      `To: ${m.to.map(who).join(', ')}`,
      m.cc?.length ? `Cc: ${m.cc.map(who).join(', ')}` : '',
      `Date: ${new Date(m.date).toUTCString()}`,
      `Subject: ${m.subject}`,
      '',
      (m.text || '').slice(0, 6000),
    ].filter((l, i) => l !== '' || i === 5).join('\n');
    if (blocks.length && total + block.length > 24000) break;
    blocks.unshift(block);
    total += block.length;
  }
  return `<email_thread>\n${blocks.join('\n\n---\n\n')}\n</email_thread>

Write my reply to the latest message from ${who(target.from)}.

<my_note>
${note || '(No note. Write a sensible reply based on the thread, with [placeholders] for anything I must fill in.)'}
</my_note>${standard ? `

<my_standard_answer>
${standard}
</my_standard_answer>
This is my usual answer to this kind of question. Follow its wording and links where it fits the email; drop what does not apply.` : ''}`;
}

// "My draft only" mode: Claude never sees the customer's email, only my note and my own text.
export function noteReplyPrompt({ first, note, standard }) {
  return `Write a reply email to ${first || 'a customer'}. You do not see their email; my note says what the reply should say.

<my_note>
${note}
</my_note>${standard ? `

<my_standard_answer>
${standard}
</my_standard_answer>
Use it as the starting point where it fits.` : ''}`;
}

export const reviseDraftPrompt = (draft, instruction, isNew) => `Here is my email draft:

<draft>
${draft}
</draft>

Revise it: ${instruction}

Return the complete revised email only${isNew ? ', starting with the "Subject: ..." line' : ''}. The same rules apply.`;

export function newEmailPrompt({ to, note }) {
  return `Write a new email from me to ${to.length ? to.map(who).join(', ') : '(recipient not given)'}. Write it in English unless my note asks for another language.

<my_note>
${note}
</my_note>

Start with a line "Subject: ..." then a blank line, then the email body.`;
}

export const revisePrompt = (instruction, isNew) =>
  `Revise the email: ${instruction}\n\nReturn the complete revised email only${isNew ? ', starting again with the "Subject: ..." line' : ''}. The same rules apply.`;

export function splitSubject(text) {
  const m = text.match(/^\s*Subject:\s*(.+)\n+/i);
  return m ? { subject: m[1].trim(), body: text.slice(m[0].length).trim() } : { subject: '', body: text.trim() };
}

// Safety net: the signature (with "Best regards") is added by Outlook or the app, so drop any the model adds.
export function tidy(text) {
  return text
    .replace(/^```[a-z]*\n?|\n?```$/g, '')
    .replace(/\n+\s*(best|kind|warm|warmest)\s+regards,?(\s*\n.{0,60}){0,3}\s*$/i, '')
    .trim();
}

async function apiError(res) {
  let msg = `Claude API error ${res.status}`;
  try { const j = await res.json(); msg = j.error?.message || msg; } catch { /* no body */ }
  if (res.status === 401) msg = 'Your Claude API key was rejected. Check it in Settings.';
  if (res.status === 429) msg = 'Claude rate limit reached. Wait a moment and try again.';
  if (res.status === 529) msg = 'Claude is overloaded right now. Try again in a minute.';
  const e = new Error(msg);
  e.status = res.status;
  return e;
}

export async function listModels(key) {
  const res = await fetch(`${API}/models?limit=100`, { headers: headers(key) });
  if (!res.ok) throw await apiError(res);
  return (await res.json()).data || [];
}

export function pickModel(models) {
  const live = models.filter((m) => !m.deprecated_at);
  const pool = live.length ? live : models;
  const newest = (line) => pool
    .filter((m) => (m.line || m.id || '').toLowerCase().includes(line))
    .sort((a, b) => new Date(b.created_at) - new Date(a.created_at))[0];
  return (newest('sonnet') || newest('opus') || pool[0])?.id || FALLBACK_MODEL;
}

export function prettyModel(id = '') {
  const parts = id.replace(/^claude-/, '').split('-').filter((p) => !/^\d{8}$/.test(p));
  const name = parts.filter((p) => !/^\d+$/.test(p)).map((p) => p[0].toUpperCase() + p.slice(1)).join(' ');
  const ver = parts.filter((p) => /^\d+$/.test(p)).join('.');
  return `Claude ${name}${ver ? ` ${ver}` : ''}`;
}

// Streams the draft; onText receives the text written so far.
export async function draft({ key, model, effort, system, messages, onText, signal }) {
  const body = { model, max_tokens: 4096, system, messages, stream: true };
  if (effort) body.output_config = { effort };
  const send = () => fetch(`${API}/messages`, { method: 'POST', headers: headers(key), body: JSON.stringify(body), signal });
  let res = await send();
  if (res.status === 400 && body.output_config) {
    const err = await apiError(res);
    if (!/output_config|effort/i.test(err.message)) throw err;
    delete body.output_config; // older model without the effort setting
    res = await send();
  }
  if (!res.ok) throw await apiError(res);

  const reader = res.body.getReader();
  const dec = new TextDecoder();
  let buf = '', text = '', stop = '', usage = {};
  for (;;) {
    const { value, done } = await reader.read();
    if (done) break;
    buf += dec.decode(value, { stream: true }).replace(/\r\n/g, '\n');
    let i;
    while ((i = buf.indexOf('\n\n')) >= 0) {
      const chunk = buf.slice(0, i);
      buf = buf.slice(i + 2);
      const data = chunk.split('\n').filter((l) => l.startsWith('data:')).map((l) => l.slice(5).trim()).join('');
      if (!data) continue;
      let ev;
      try { ev = JSON.parse(data); } catch { continue; }
      if (ev.type === 'content_block_delta' && ev.delta?.type === 'text_delta') {
        text += ev.delta.text;
        onText?.(text);
      } else if (ev.type === 'message_start') {
        usage = { ...(ev.message?.usage || {}) };
      } else if (ev.type === 'message_delta') {
        stop = ev.delta?.stop_reason || stop;
        usage = { ...usage, ...(ev.usage || {}) };
      } else if (ev.type === 'error') {
        throw new Error(ev.error?.message || 'Claude stopped with an error.');
      }
    }
  }
  if (!text.trim()) throw new Error(stop === 'max_tokens' ? 'Claude ran out of room before writing. Try again or pick "Fast" in Settings.' : 'Claude returned an empty draft. Try again.');
  return { text: tidy(text), usage, stop };
}

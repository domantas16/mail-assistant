// Demo mode: fictional sample mailbox so the app can be tried before Microsoft 365 is connected.
import { groupChats } from './lib.js';

const H = 3600e3, D = 24 * H;
const ago = (ms) => new Date(Date.now() - ms).toISOString();
const p = (name, email, company = '', jobTitle = '') => ({ name, email, company, jobTitle, phones: [] });

const ME = { name: 'Demo User', email: 'me@optoman.example' };
const PEOPLE = {
  elena: p('Dr. Elena Varga', 'e.varga@photonix-labs.example', 'Photonix Labs', 'Senior Laser Engineer'),
  li: p('Li Wei', 'li.wei@lightwave-sz.example', 'Lightwave Shenzhen', 'Purchasing'),
  marco: p('Marco Bianchi', 'm.bianchi@ottica-nova.example', 'Ottica Nova S.r.l.', 'Procurement'),
  anna: p('Prof. Anna Lind', 'anna.lind@nordlicht-institute.example', 'Nordlicht Institute', 'Professor of Physics'),
  kenji: p('Kenji Sato', 'k.sato@hikari-trading.example', 'Hikari Optics Trading', 'Sales Director'),
  expo: p('Photonics Expo Team', 'events@expo-news.example', 'Photonics Expo'),
  rasa: p('Rasa P.', 'rasa@optoman.example', 'OPTOMAN', 'Marketing'),
};
PEOPLE.elena.phones = ['+36 1 555 0142'];
PEOPLE.kenji.phones = ['+81 3 5550 1180'];

const para = (...ps) => ps.map((t) => `<p>${t}</p>`).join('');
let n = 0;
const msg = (conv, from, to, when, subject, html, extra = {}) => ({
  id: `demo-${++n}`, conversationId: conv, from, to, cc: extra.cc || [], date: ago(when), subject, bodyHtml: html,
  preview: html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 140),
  isRead: extra.isRead ?? true, isDraft: false, hasAttachments: !!extra.atts, atts: extra.atts || [], flagged: !!extra.flagged,
  important: false, webLink: null, folderId: extra.folder || 'inbox',
});

function sample() {
  n = 0;
  const { elena, li, marco, anna, kenji, expo, rasa } = PEOPLE;
  return [
    msg('c0', elena, [ME], 12 * D, 'Samples for evaluation?', para(
      'Hi,', 'Do you offer sample pieces of your low-GDD mirrors, so we can evaluate them before a larger order?', 'Elena',
    )),
    msg('c0', ME, [elena], 11.8 * D, 'RE: Samples for evaluation?', para(
      'Dear Elena,', "Great to hear from you! We don't provide free samples, but you can order single pieces directly in OPTOSHOP, with no minimum quantity.", 'Keep in touch!',
    ), { folder: 'sentitems' }),
    msg('c1', elena, [ME], 2.5 * H, 'Quote request: HR mirrors for 1030 nm', para(
      'Dear OPTOMAN team,',
      'We are building a new Yb:KGW amplifier and need high-reflectivity mirrors for 1030 nm at 45° AOI, 25.4 mm diameter, p-polarization.',
      'Could you quote 10 and 20 pieces? We would also like to know the LIDT for ~300 fs pulses and your current lead time.',
      'Best regards,<br>Elena Varga<br>Senior Laser Engineer, Photonix Labs',
    ), { isRead: false, cc: [p('Tamas Kovacs', 't.kovacs@photonix-labs.example')], atts: [{ id: 'a1', name: 'beamline-layout.pdf', type: 'application/pdf', size: 482113, inline: false }] }),
    msg('c2', li, [ME], 4 * H, 'Shipping to Shenzhen - customs documents', para(
      'Hello,',
      'We want to order thin-film polarizers from OPTOSHOP, but last time another supplier\'s parcel was held by customs for two weeks.',
      'Which courier do you use, and do you provide a commercial invoice with the correct HS code?',
      'Thanks,<br>Li Wei',
    ), { isRead: false }),
    msg('c3', marco, [ME], 3 * D, 'Order OS-10422 - delivery date', para(
      'Hello,', 'We placed order OS-10422 on OPTOSHOP for 5x thin-film polarizers. Could you confirm the delivery date?', 'Regards,<br>Marco',
    )),
    msg('c3', ME, [marco], 2 * D, 'RE: Order OS-10422 - delivery date', para(
      'Dear Marco,', 'Thank you for your order! The polarizers ship this Friday via DHL, and you will receive the tracking number by email.', 'Keep in touch!',
    ), { folder: 'sentitems' }),
    msg('c3', marco, [ME], 20 * H, 'RE: Order OS-10422 - delivery date', para(
      'Thanks, great news!', 'Could you also send the test reports for this batch? Our QA team needs them before installation.', 'Marco',
    ) + '<blockquote><p>Dear Marco, Thank you for your order! The polarizers ship this Friday via DHL...</p></blockquote>', { isRead: false }),
    msg('c4', anna, [ME], 1.2 * D, 'Academic discount for our teaching lab?', para(
      'Dear Sir or Madam,',
      'We are setting up an ultrafast optics teaching lab for master students and are looking at your low-GDD mirrors on OPTOSHOP.',
      'Do you offer any discount for universities? We would need about 12 mirrors.',
      'Kind regards,<br>Anna Lind',
    )),
    msg('c5', kenji, [ME], 2 * D, 'Q4 stock update and forecast', para(
      'Hi,', 'Please find our Q4 forecast below. We expect demand for 1064 nm mirrors to grow about 20% next quarter.',
      'Could you confirm that you can hold 50 pcs of HR1064-25 in stock for us until January?', 'Best,<br>Kenji',
    ), { flagged: true }),
    msg('c6', expo, [ME], 4 * D, 'Photonics Expo 2027 - exhibitor information',
      '<div style="font-family:Arial;max-width:560px"><img src="https://images.expo-news.example/banner.png" alt="Photonics Expo banner" width="560" height="120"><h2 style="color:#123a7a">Book your booth for 2027</h2><p>Early-bird rates end on 15 January. Exhibitors get a free listing in the digital catalogue.</p><p><a href="https://expo-news.example/book">Book now</a></p></div>'),
    msg('c7', rasa, [ME], 5 * D, 'Lunch on Friday?', para('Hi! Are you joining the team lunch on Friday at 12:30? 🙂', 'Rasa')),
  ];
}

export function cannedDraft({ target, note, isNew }) {
  const first = (target?.from?.name || '').replace(/^(dr|prof)\.?\s+/i, '').split(' ')[0] || 'there';
  const body = `Dear ${first},\n\nThank you for your message!${note ? ` [Your note "${note}" would be turned into a polished paragraph here.]` : ''}\n\nThis is a sample draft from demo mode. Add your Claude API key in Settings to generate real drafts from your short notes.\n\nKeep in touch!`;
  return isNew ? `Subject: [Subject from your note]\n\n${body}` : body;
}

export const DEMO_SHARED = 'info@optoman.example';
const INFO = { name: 'OPTOMAN Info', email: DEMO_SHARED };

function sharedSample() {
  const tomas = p('Tomás Ruiz', 't.ruiz@laserworks.example', 'LaserWorks Iberia', 'Lab Manager');
  const yuki = p('Yuki Tanaka', 'y.tanaka@kotobuki-optics.example', 'Kotobuki Optics', 'Buyer');
  return [
    msg('s1', tomas, [INFO], 3 * H, 'Catalogue request', para('Hello,', 'Could you send us your latest product catalogue? We are planning purchases for Q1.', 'Thanks,<br>Tomás'), { isRead: false }),
    msg('s2', yuki, [INFO], 1.5 * D, 'HTS code for mirrors', para('Dear OPTOMAN,', 'For customs we need the HTS code and ECCN for your dielectric mirrors. Could you confirm?', 'Best regards,<br>Yuki Tanaka')),
  ];
}

export function createDemo() {
  const boxes = { me: { all: sample(), drafts: [], self: ME }, [DEMO_SHARED]: { all: sharedSample(), drafts: [], self: INFO } };
  let box = boxes.me;
  let all = box.all;
  let drafts = box.drafts;
  const byFolder = (f) => (f === 'drafts' ? drafts : all.filter((m) => m.folderId === f && !m.isDraft));
  const latestPerConv = (list) => list.slice().sort((a, b) => new Date(b.date) - new Date(a.date));
  const people = Object.values(PEOPLE).filter((x) => x !== PEOPLE.expo);
  const match = (x, q) => `${x.name} ${x.email} ${x.company}`.toLowerCase().includes(q.toLowerCase());
  const later = (v) => new Promise((r) => setTimeout(() => r(v), 150));

  return {
    demo: true,
    setMailbox(address) {
      box = boxes[address] || boxes.me;
      all = box.all;
      drafts = box.drafts;
    },
    send(draftId) {
      const i = drafts.findIndex((d) => d.id === draftId);
      if (i >= 0) {
        const [d] = drafts.splice(i, 1);
        all.push({ ...d, id: `sent-${Date.now()}`, isDraft: false, folderId: 'sentitems', from: box.self, date: new Date().toISOString() });
      }
      return later();
    },
    me: () => later({ ...ME, jobTitle: '' }),
    folders: () => later({
      main: [
        { id: 'inbox', key: 'inbox', name: 'Inbox', icon: 'inbox', unread: all.filter((m) => m.folderId === 'inbox' && !m.isRead).length },
        { id: 'drafts', key: 'drafts', name: 'Drafts', icon: 'draft', unread: 0 },
        { id: 'sentitems', key: 'sentitems', name: 'Sent', icon: 'send', unread: 0 },
        { id: 'archive', key: 'archive', name: 'Archive', icon: 'archive', unread: 0 },
      ],
      other: box === boxes.me ? [{ id: 'customers', name: 'Customers', icon: 'folder', unread: 0, child: true }] : [],
    }),
    messages(folderId, { search } = {}) {
      let list;
      if (search) {
        const q = search.replace(/^participants:/i, '');
        list = all.filter((m) => [m.from, ...m.to].some((x) => match(x, q)) || `${m.subject} ${m.preview}`.toLowerCase().includes(q.toLowerCase()));
      } else list = byFolder(folderId);
      return later({ items: latestPerConv(list), next: null });
    },
    thread: (m) => later(all.filter((x) => x.conversationId === m.conversationId).sort((a, b) => new Date(a.date) - new Date(b.date))),
    fullBody: (id) => later(all.find((m) => m.id === id)?.bodyHtml || ''),
    attachments: (id) => later(all.find((m) => m.id === id)?.atts || []),
    inlineImages: () => later({}),
    attachmentBlob: (id, aid) => later(new Blob([`Demo attachment ${aid}`], { type: 'text/plain' })),
    markRead(id, isRead = true) { const m = all.find((x) => x.id === id); if (m) m.isRead = isRead; return later(); },
    createReplyDraft(id, { html }) {
      const src = all.find((m) => m.id === id);
      drafts.unshift({ ...src, id: `draft-${Date.now()}`, from: box.self, to: [src.from], cc: [], subject: `RE: ${src.subject.replace(/^RE:\s*/i, '')}`, bodyHtml: html, preview: html.replace(/<[^>]+>/g, ' ').trim().slice(0, 140), isDraft: true, isRead: true, date: new Date().toISOString(), folderId: 'drafts', conversationId: src.conversationId });
      return later({ id: drafts[0].id, webLink: null });
    },
    createDraft({ to, subject, html }) {
      drafts.unshift({ id: `draft-${Date.now()}`, conversationId: `d${Date.now()}`, from: box.self, to, cc: [], subject: subject || '(no subject)', bodyHtml: html, preview: html.replace(/<[^>]+>/g, ' ').trim().slice(0, 140), isDraft: true, isRead: true, date: new Date().toISOString(), folderId: 'drafts', atts: [] });
      return later({ id: drafts[0].id, webLink: null });
    },
    chats: () => later(groupChats(all, box.self.email)),
    chat: (email) => later(all.filter((m) => [m.from, ...m.to, ...m.cc].some((x) => x?.email === email)).sort((a, b) => new Date(a.date) - new Date(b.date))),
    people: (q = '') => later(q ? people.filter((x) => match(x, q)) : people),
    contacts: () => later(people.slice().sort((a, b) => a.name.localeCompare(b.name))),
    sentExamples: () => later([]),
  };
}

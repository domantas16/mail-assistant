// Standard answers: ready replies picked locally by keywords, so a draft is waiting
// even with AI switched off (nothing leaves Microsoft and this browser).
// Format, editable in Settings:
//   ## Name
//   keywords: word, another phrase, ...
//   Body text. {first} becomes the recipient's first name. [square brackets] mark what to fill in.

export const DEFAULT_ANSWERS = `## Quote request
keywords: quote, quotation, price, pricing, offer, rfq, cost
Dear {first},

Great to hear from you! For [product and quantity], the price is [price] per piece, and the current lead time is about [lead time].

You can order, or request for a quotation directly in OPTOSHOP: https://optoshop.optoman.com

Keep in touch!

## Order status and delivery
keywords: order, delivery, tracking, shipped, ship date, dispatch, delivery date
Dear {first},

Thank you for your order! It ships on [date] via [DHL / FedEx], and you will receive the tracking number by email.

Keep in touch!

## Documents and test reports
keywords: test report, test reports, certificate, measurement report, qa, documentation, datasheet
Dear {first},

Thank you for your message! [Answer about the requested documents.]

Let me know if you have any questions.

## Shipping and customs
keywords: customs, courier, shipping, hs code, import, commercial invoice, dhl, fedex
Dear {first},

Thank you for reaching out! Our goods ship directly from Lithuania via FedEx or DHL, and customs clearance has generally been smooth as long as the consignee meets the import requirements. We declare the correct commercial value and provide the commercial invoice and all shipping documents, and we can supply additional documents promptly on request.

Let me know if you have any questions.

## HTS and ECCN
keywords: hts, eccn, export control, classification, tariff code
Dear {first},

Thank you for your message! Our standard classification is HTS code 9001.90.9000 and ECCN EAR99.

Let me know if you have any questions.

## Custom optics or technical inquiry
keywords: custom, custom coating, custom optics, specification, specifications, feasibility
Dear {first},

You came to the right place! The quickest way to find the right solution is our OPTOguide: https://www.optoman.com/optoguide/

Or simply send us your specifications: wavelength, AOI, polarization, fluence or power, pulse width, repetition rate, beam size and quantity, and we will prepare a quotation.

Keep in touch!

## Catalogue request
keywords: catalogue, catalog, brochure, product list
Dear {first},

Thank you for your interest! OPTOMAN is not a catalogue company, and standard products are only a small part of what we make. The up-to-date list of standard products is always in OPTOSHOP: https://optoshop.optoman.com

Keep in touch!

## Academic discount
keywords: discount, academic, university, universities, student, students, teaching
Dear {first},

Great to hear from you! [Discount details for your institution.]

You can order, or request for a quotation directly in OPTOSHOP: https://optoshop.optoman.com

Keep in touch!`;

export const GENERIC = {
  name: 'General reply',
  keywords: [],
  body: 'Dear {first},\n\nThank you for your message! [Your answer.]\n\nKeep in touch!',
};

let cacheText = null;
let cacheList = [];

export function parseAnswers(text) {
  if (text === cacheText) return cacheList;
  const list = [];
  for (const block of String(text || '').split(/^##\s+/m).slice(1)) {
    const lines = block.split('\n');
    const name = lines.shift().trim();
    let keywords = [];
    if (/^\s*keywords\s*:/i.test(lines[0] || '')) {
      keywords = lines.shift().replace(/^\s*keywords\s*:/i, '').split(',').map((k) => k.trim().toLowerCase()).filter(Boolean);
    }
    const body = lines.join('\n').trim();
    if (name && body) list.push({ name, keywords, body });
  }
  cacheText = text;
  cacheList = list;
  return list;
}

const hits = (hay, k) => {
  const re = new RegExp(`(^|[^a-z0-9])${k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}($|[^a-z0-9])`, 'i');
  return re.test(hay);
};

// Best standard answer for an email, or null when nothing fits. Words in the latest message
// count double; the subject (often an old thread title) counts once.
export function pickAnswer(list, { subject = '', text = '' }) {
  const subj = subject.toLowerCase();
  const body = text.toLowerCase();
  let best = null, top = 0;
  for (const a of list) {
    let score = 0;
    for (const k of a.keywords) {
      if (hits(body, k)) score += 2;
      else if (hits(subj, k)) score += 1;
    }
    if (score > top) { best = a; top = score; }
  }
  return best;
}

export function firstName(person) {
  const name = (person?.name || '').trim();
  if (!name || name.includes('@')) return '';
  const parts = name.replace(/^(dr|prof|mr|mrs|ms)\.?\s+/i, '').split(/\s+/);
  return parts[0] || '';
}

export function fillAnswer(answer, person) {
  const first = firstName(person);
  return answer.body
    .replace(/\{first\}/g, first || '[name]')
    .trim();
}

import domino from '@mixmark-io/domino';

/**
 * Turn a newsletter-style email body into readable plain text plus the links
 * a reader would actually click. Pure: takes the JMAP Email object (with
 * htmlBody/textBody/bodyValues) and returns data.
 *
 * `text` carries a " [n]" marker after every kept link's anchor text, where n
 * is the link's 1-based index in `links`. That keeps "which link belongs to
 * which item" unambiguous when many anchors share a label (e.g. every puzzle
 * in Artisanal Sudoku links "SudokuPad").
 */

export interface EmailLink {
  text: string;
  url: string;
}

export interface EmailContent {
  text: string;
  links: EmailLink[];
  webVersionUrl: string | null;
}

// Characters used as invisible preheader padding or soft breaks.
const INVISIBLE_CHARS = /[\u034F\u00AD\u180E\u200B-\u200F\u2028\u2029\u2060-\u2064\uFEFF]/g;
const NBSP_LIKE = /[\u00A0\u2007\u202F]/g;

const DROP_TAGS = new Set([
  'head', 'style', 'script', 'noscript', 'template', 'title', 'meta', 'link',
  'img', 'picture', 'svg', 'video', 'audio', 'iframe', 'object', 'map', 'button', 'input', 'select',
]);

const BLOCK_TAGS = new Set([
  'address', 'article', 'aside', 'blockquote', 'center', 'dd', 'div', 'dl', 'dt', 'fieldset',
  'figcaption', 'figure', 'footer', 'form', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'header', 'hr',
  'li', 'main', 'nav', 'ol', 'p', 'pre', 'section', 'table', 'tbody', 'thead', 'tfoot', 'tr',
  'td', 'th', 'ul',
]);

const WEB_VERSION_TEXT =
  /\b(view|read|open|see)\b[^.]{0,40}?\b(in (your |a )?(web )?browser|online|on (the )?web|in (the )?app|web version)\b|^web version$|no images\?/i;
const GENERIC_ANCHOR_TEXT = /^(click here|here|click)$/i;

const BOILERPLATE_TEXT =
  /^(unsubscribe.*|manage (your )?(email |subscription )?preferences|update (your )?(email )?preferences|email preferences|subscribe( free| here| now| for free)?|sign up.*|start (a |free |your )?trial|like|comment|restack|share|forward|tweet|facebook|x|twitter|instagram|bluesky|threads|youtube|tiktok|linkedin|mastodon|more like this|less like this|email|get the app|download the app|powered by .*)$/i;

const BOILERPLATE_HREF =
  /unsubscribe|\/members\/feedback\/|disable_email|[?&]submitLike=|[?&]comments=true|utm_content=share|email-restack|substack\.com\/signup|ghost\.org\/\?via|list-manage\.com\/(un)?subscribe/i;

const SUBSTACK_POST = /^https:\/\/open\.substack\.com\/pub\/[^/?#]+\/p\/[^/?#]+/i;

export function extractEmailContent(email: any): EmailContent {
  const bodyValues = email?.bodyValues ?? {};
  const htmlPart = email?.htmlBody?.[0]?.partId;
  const html = htmlPart ? bodyValues[htmlPart]?.value : undefined;
  if (html) return fromHtml(html);

  const textPart = email?.textBody?.[0]?.partId;
  const plain = textPart ? bodyValues[textPart]?.value : undefined;
  return fromPlainText(plain ?? '');
}

// ---------- HTML ----------

function fromHtml(html: string): EmailContent {
  const doc = domino.createDocument(html);
  const anchors = Array.from(doc.querySelectorAll('a[href]')) as Element[];

  const webVersionUrl = findWebVersion(anchors);
  const links: EmailLink[] = [];
  const linkIndex = new Map<string, number>();
  let out = '';

  const walk = (node: Node): void => {
    if (node.nodeType === 3) {
      out += (node as Text).data;
      return;
    }
    if (node.nodeType !== 1) return;
    const el = node as Element;
    const tag = el.tagName.toLowerCase();
    if (DROP_TAGS.has(tag) || isHidden(el)) return;
    if (tag === 'br') {
      out += '\n';
      return;
    }

    const block = BLOCK_TAGS.has(tag);
    if (block) out += '\n';
    for (const child of Array.from(el.childNodes)) walk(child);

    if (tag === 'a') {
      const link = keptLink(el, webVersionUrl);
      if (link) {
        const key = `${link.text}\u0000${link.url}`;
        let n = linkIndex.get(key);
        if (n === undefined) {
          links.push(link);
          n = links.length;
          linkIndex.set(key, n);
        }
        out += ` [${n}]`;
      }
    }
    if (block) out += '\n';
  };

  walk(doc.body ?? doc.documentElement);
  return { text: cleanText(out), links, webVersionUrl };
}

function isHidden(el: Element): boolean {
  const style = el.getAttribute('style') ?? '';
  return /display\s*:\s*none|visibility\s*:\s*hidden|mso-hide\s*:\s*all/i.test(style) || el.hasAttribute('hidden');
}

function anchorText(el: Element): string {
  return cleanInline(el.textContent ?? '');
}

function isWebVersionAnchor(el: Element): boolean {
  const text = anchorText(el);
  if (WEB_VERSION_TEXT.test(text)) return true;
  if (GENERIC_ANCHOR_TEXT.test(text) && el.parentElement) {
    const context = cleanInline(el.parentElement.textContent ?? '');
    return context.length <= 120 && WEB_VERSION_TEXT.test(context);
  }
  return false;
}

function findWebVersion(anchors: Element[]): string | null {
  // Substack: the canonical post URL is the web version.
  for (const a of anchors) {
    const m = SUBSTACK_POST.exec(a.getAttribute('href')?.trim() ?? '');
    if (m) return m[0];
  }
  for (const a of anchors) {
    const href = a.getAttribute('href')?.trim() ?? '';
    if (/^https?:\/\//i.test(href) && isWebVersionAnchor(a)) return href;
  }
  return null;
}

function keptLink(el: Element, webVersionUrl: string | null): EmailLink | null {
  const url = el.getAttribute('href')?.trim() ?? '';
  if (!/^https?:\/\//i.test(url)) return null;
  const text = anchorText(el);
  if (!text) return null; // image-only / tracking
  if (BOILERPLATE_TEXT.test(text) || BOILERPLATE_HREF.test(url)) return null;
  if (isWebVersionAnchor(el)) return null;
  if (webVersionUrl && url.startsWith(webVersionUrl)) return null;
  return { text, url };
}

// ---------- plain text ----------

function fromPlainText(plain: string): EmailContent {
  let webVersionUrl: string | null = null;
  for (const line of plain.split('\n')) {
    const url = /(https?:\/\/\S+)/.exec(line)?.[1];
    if (url && WEB_VERSION_TEXT.test(line)) {
      webVersionUrl = url;
      break;
    }
  }
  const links: EmailLink[] = [];
  const seen = new Set<string>();
  for (const [url] of plain.matchAll(/https?:\/\/[^\s\]]+/g)) {
    if (url === webVersionUrl || seen.has(url) || BOILERPLATE_HREF.test(url)) continue;
    seen.add(url);
    links.push({ text: url, url });
  }
  return { text: cleanText(plain), links, webVersionUrl };
}

// ---------- cleanup ----------

function cleanInline(s: string): string {
  return s.replace(INVISIBLE_CHARS, '').replace(NBSP_LIKE, ' ').replace(/\s+/g, ' ').trim();
}

/** Remove CSS that leaked into text: @media/@font-face/... blocks and bare `selector { prop: value }` rules. */
function stripLeakedCss(s: string): string {
  let result = '';
  let i = 0;
  const atRule = /@(media|font-face|import|supports|keyframes|-webkit-keyframes|page)\b/gi;
  while (i < s.length) {
    atRule.lastIndex = i;
    const m = atRule.exec(s);
    if (!m) {
      result += s.slice(i);
      break;
    }
    result += s.slice(i, m.index);
    const open = s.indexOf('{', m.index);
    if (open === -1) {
      result += s.slice(m.index);
      break;
    }
    let depth = 0;
    let j = open;
    for (; j < s.length; j++) {
      if (s[j] === '{') depth++;
      else if (s[j] === '}' && --depth === 0) break;
    }
    i = j + 1;
  }
  return result.replace(/[^\n{}]{1,200}\{[^{}]*:[^{}]*\}/g, '');
}

function cleanText(raw: string): string {
  const text = stripLeakedCss(raw.replace(INVISIBLE_CHARS, '').replace(NBSP_LIKE, ' '));
  const lines = text.split(/\r?\n/).map(line => line.replace(/[ \t\f\v]+/g, ' ').trim());
  return lines
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

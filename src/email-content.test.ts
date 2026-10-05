import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'fs';
import { extractEmailContent } from './email-content.js';

function fixture(name: string) {
  return JSON.parse(readFileSync(new URL(`./fixtures/${name}.json`, import.meta.url), 'utf8'));
}

const INVISIBLE = /[\u034F\u00AD\u200B-\u200F\u2060\uFEFF\u00A0]/;
const BOILERPLATE_TEXT = /^(unsubscribe|like|comment|restack|facebook|x|instagram|bluesky|threads|youtube|tiktok|more like this|less like this|subscribe|subscribe free|start trial|manage your email preferences)$/i;

function assertCleanText(text: string) {
  assert.ok(text.length > 200, 'text should have content');
  assert.doesNotMatch(text, /@media/);
  assert.doesNotMatch(text, /\{[^}]*:[^}]*\}/, 'no CSS rule bodies');
  assert.doesNotMatch(text, INVISIBLE, 'no invisible padding characters');
  assert.doesNotMatch(text, /[ \t]{2,}/, 'no runs of spaces');
  assert.doesNotMatch(text, /\n{3,}/, 'no runs of blank lines');
  assert.doesNotMatch(text, /<img|<style|<\/?(div|td|table)/i, 'no markup');
}

function assertNoBoilerplateLinks(links: { text: string; url: string }[]) {
  for (const l of links) {
    assert.ok(l.text.trim().length > 0, `image-only link leaked: ${l.url}`);
    assert.doesNotMatch(l.text, BOILERPLATE_TEXT, `boilerplate link leaked: ${l.text}`);
    assert.match(l.url, /^https?:\/\//);
  }
  const keys = links.map(l => `${l.text}\u0000${l.url}`);
  assert.equal(new Set(keys).size, keys.length, 'links are deduped');
}

describe('extractEmailContent — Artisanal Sudoku (Substack)', () => {
  const content = extractEmailContent(fixture('artisanal-sudoku-247'));

  it('produces clean readable text containing the three puzzle names', () => {
    assertCleanText(content.text);
    for (const name of ['Melon Baller', 'Get the Horns', 'Water Lilies']) {
      assert.match(content.text, new RegExp(name));
    }
  });

  it('includes the four "other puzzles I enjoyed" links with their text', () => {
    for (const name of ['Quincinx', '159-X', 'Killers in the dark', 'This Hits Different']) {
      const link = content.links.find(l => l.text === name);
      assert.ok(link, `missing link for ${name}`);
      assert.match(link!.url, /^https:\/\/substack\.com\/redirect\//);
    }
  });

  it('ties each puzzle to its own SudokuPad link via [n] markers in the text', () => {
    const urls = new Set<string>();
    for (const [name, next] of [['Melon Baller', 'Get the Horns'], ['Get the Horns', 'Water Lilies'], ['Water Lilies', 'Some other puzzles']]) {
      const section = content.text.slice(content.text.indexOf(name), content.text.indexOf(next));
      const m = /SudokuPad \[(\d+)\]/.exec(section);
      assert.ok(m, `no SudokuPad marker in section for ${name}`);
      const link = content.links[Number(m![1]) - 1];
      assert.equal(link.text, 'SudokuPad');
      urls.add(link.url);
    }
    assert.equal(urls.size, 3, 'three distinct puzzle links');
  });

  it('marks links in document order', () => {
    const indices = [...content.text.matchAll(/\[(\d+)\]/g)].map(m => Number(m[1]));
    const firstSeen = indices.filter((n, i) => indices.indexOf(n) === i);
    assert.deepEqual(firstSeen, firstSeen.slice().sort((a, b) => a - b));
    assert.equal(Math.max(...indices), content.links.length);
  });

  it('excludes image-only, share, unsubscribe and social links', () => {
    assertNoBoilerplateLinks(content.links);
  });

  it('uses the canonical Substack post URL as the web version', () => {
    assert.equal(content.webVersionUrl, 'https://open.substack.com/pub/artisanalsudoku/p/artisanal-sudoku-volume-247');
  });
});

describe('extractEmailContent — Killscreen digest', () => {
  const content = extractEmailContent(fixture('killscreen-digest'));

  it('produces clean readable text', () => {
    assertCleanText(content.text);
    assert.match(content.text, /Murlo Builds a World You Can Play/);
  });

  it('keeps article links with their titles', () => {
    for (const title of ['Murlo Builds a World You Can Play', 'Playable poems now come with peer review', 'Lofsöng hides a nuclear warning']) {
      assert.ok(content.links.some(l => l.text === title), `missing ${title}`);
    }
    assertNoBoilerplateLinks(content.links);
  });

  it('detects "View in browser" as the web version', () => {
    assert.equal(content.webVersionUrl, 'https://www.killscreen.com/r/bf00627b?m=REDACTED');
    assert.ok(!content.links.some(l => /view in browser/i.test(l.text)), 'web-version link not repeated in links');
  });
});

describe('extractEmailContent — The Saturday Paper', () => {
  const content = extractEmailContent(fixture('saturday-paper-arts'));

  it('produces clean readable text', () => {
    assertCleanText(content.text);
    assert.match(content.text, /Cartoonish Cruise ends up in a ham-fisted hole in Digger/);
  });

  it('detects "No images? Click here" as the web version', () => {
    assert.equal(content.webVersionUrl, 'https://campaigns.schwartzmedia.com.au/t/i-e-aaaaaaa-bbbbbbbbbb-ii/');
  });

  it('keeps story links and drops social/footer boilerplate', () => {
    assert.ok(content.links.some(l => l.text === 'Cartoonish Cruise ends up in a ham-fisted hole in Digger'));
    assertNoBoilerplateLinks(content.links);
  });
});

describe('extractEmailContent — edge cases', () => {
  function htmlEmail(html: string) {
    return { htmlBody: [{ partId: '1' }], bodyValues: { '1': { value: html } } };
  }

  it('strips leaked @media CSS text and tracking pixels', () => {
    const c = extractEmailContent(htmlEmail(
      '<div>@media only screen and (max-width: 600px) { .x { width: 100% !important; } }</div>' +
      '<p>Hello\u034F\u00AD\u200B&nbsp;&nbsp;&nbsp;world</p><img src="https://t.example/pixel.gif" width="1" height="1">',
    ));
    assert.equal(c.text, 'Hello world');
    assert.deepEqual(c.links, []);
    assert.equal(c.webVersionUrl, null);
  });

  it('dedupes exact duplicate links but keeps same text with different urls', () => {
    const c = extractEmailContent(htmlEmail(
      '<p><a href="https://a.example/1">Story</a> and <a href="https://a.example/1">Story</a> and <a href="https://a.example/2">Story</a></p>',
    ));
    assert.deepEqual(c.links, [
      { text: 'Story', url: 'https://a.example/1' },
      { text: 'Story', url: 'https://a.example/2' },
    ]);
    assert.equal(c.text, 'Story [1] and Story [1] and Story [2]');
  });

  it('falls back to the plain-text body when there is no HTML', () => {
    const c = extractEmailContent({
      textBody: [{ partId: '1' }],
      bodyValues: { '1': { value: 'View this post on the web at https://x.substack.com/p/slug\n\nHi   there\u200B' } },
    });
    assert.equal(c.webVersionUrl, 'https://x.substack.com/p/slug');
    assert.match(c.text, /Hi there$/);
  });
});

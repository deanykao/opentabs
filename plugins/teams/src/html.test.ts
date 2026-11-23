import { describe, expect, test } from 'vitest';
import { buildMessageBody, escapeHtml, randomClientMessageId, textToHtml } from './html.js';

describe('textToHtml', () => {
  test('escapes markup and converts newlines', () => {
    expect(textToHtml('a < b && c > d\nline2\r\nline3')).toBe('a &lt; b &amp;&amp; c &gt; d<br>line2<br>line3');
  });
  test('escapes ampersand before angle brackets (no double escaping of entities it produced)', () => {
    expect(escapeHtml('<&>')).toBe('&lt;&amp;&gt;');
  });
});

describe('buildMessageBody', () => {
  test('text mode without mentions escapes', () => {
    expect(buildMessageBody('<b>hi</b>', 'text')).toEqual({ content: '&lt;b&gt;hi&lt;/b&gt;' });
  });
  test('html mode without mentions passes through', () => {
    expect(buildMessageBody('<b>hi</b>', 'html')).toEqual({ content: '<b>hi</b>' });
  });
  test('text mode replaces @Name with a Mention span and emits properties.mentions', () => {
    const { content, mentionsJson } = buildMessageBody('Hi @Ann & @Bob!', 'text', [
      { mri: '8:orgid:aaa', display_name: 'Ann' },
      { mri: '8:orgid:bbb', display_name: 'Bob' },
    ]);
    expect(content).toBe(
      'Hi <span itemscope itemtype="http://schema.skype.com/Mention" itemid="0">Ann</span> &amp; <span itemscope itemtype="http://schema.skype.com/Mention" itemid="1">Bob</span>!',
    );
    expect(JSON.parse(mentionsJson ?? '')).toEqual([
      {
        '@type': 'http://schema.skype.com/Mention',
        itemid: 0,
        mri: '8:orgid:aaa',
        mentionType: 'person',
        displayName: 'Ann',
      },
      {
        '@type': 'http://schema.skype.com/Mention',
        itemid: 1,
        mri: '8:orgid:bbb',
        mentionType: 'person',
        displayName: 'Bob',
      },
    ]);
  });
  test('display names are escaped inside the span', () => {
    const { content } = buildMessageBody('cc @A&B', 'text', [{ mri: '8:orgid:x', display_name: 'A&B' }]);
    expect(content).toContain('>A&amp;B</span>');
  });
  test('html mode uses {{n}} placeholders', () => {
    const { content } = buildMessageBody('<p>{{0}} hi</p>', 'html', [{ mri: '8:orgid:aaa', display_name: 'Ann' }]);
    expect(content).toBe('<p><span itemscope itemtype="http://schema.skype.com/Mention" itemid="0">Ann</span> hi</p>');
  });
  test('channel MRIs get mentionType channel', () => {
    const { mentionsJson } = buildMessageBody('@Chan', 'text', [{ mri: '19:abc@thread.tacv2', display_name: 'Chan' }]);
    expect(JSON.parse(mentionsJson ?? '')[0].mentionType).toBe('channel');
  });
  test('a mention that is not placed is an error, not silently dropped', () => {
    expect(() => buildMessageBody('no names here', 'text', [{ mri: '8:orgid:a', display_name: 'Ann' }])).toThrow(
      /not placed/,
    );
    expect(() => buildMessageBody('{{3}}', 'html', [{ mri: '8:orgid:a', display_name: 'Ann' }])).toThrow(
      /no matching mention/,
    );
  });
});

describe('randomClientMessageId', () => {
  test('is numeric and varies', () => {
    const a = randomClientMessageId();
    const b = randomClientMessageId();
    expect(a).toMatch(/^\d{1,20}$/);
    expect(a).not.toBe(b);
  });
});

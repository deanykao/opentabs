import { ToolError } from '@opentabs-dev/plugin-sdk';

export interface Mention {
  mri: string;
  display_name: string;
}

/** Escape text for inclusion in HTML element content. */
export const escapeHtml = (text: string): string =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** Convert plain text to Teams-safe HTML: escape &,<,> and turn newlines into <br>. */
export const textToHtml = (text: string): string => escapeHtml(text).replace(/\r\n|\r|\n/g, '<br>');

export const MENTION_TYPE = 'http://schema.skype.com/Mention';

/**
 * Build the message body and `properties.mentions` payload Teams web sends for
 * @mentions. Each mention is rendered as
 * `<span itemscope itemtype="http://schema.skype.com/Mention" itemid="N">Name</span>`
 * and described in the JSON array stored in `properties.mentions`.
 *
 * Text mode: every `@{display_name}` occurrence of a mention is replaced by its
 * span (first unused occurrence per mention; mentions are matched in order).
 * HTML mode: callers write `{{N}}` (zero-based mention index) where the span goes.
 * A mention that is never placed is an error rather than a silently dropped mention.
 */
export const buildMessageBody = (
  text: string,
  format: 'text' | 'html',
  mentions: readonly Mention[] = [],
): { content: string; mentionsJson?: string } => {
  const spans = mentions.map(
    (m, i) => `<span itemscope itemtype="${MENTION_TYPE}" itemid="${String(i)}">${escapeHtml(m.display_name)}</span>`,
  );

  if (mentions.length === 0) {
    return { content: format === 'html' ? text : textToHtml(text) };
  }

  let content: string;
  const placed = new Set<number>();
  if (format === 'html') {
    content = text.replace(/\{\{(\d+)\}\}/g, (match, idx: string) => {
      const i = Number(idx);
      const span = spans[i];
      if (span === undefined) throw ToolError.validation(`Mention placeholder ${match} has no matching mention`);
      placed.add(i);
      return span;
    });
  } else {
    // Substitute on escaped text; the escaped display name is what appears in the escaped body.
    content = textToHtml(text);
    mentions.forEach((m, i) => {
      const token = `@${escapeHtml(m.display_name)}`;
      const at = content.indexOf(token);
      if (at < 0) return;
      const span = spans[i];
      if (span === undefined) return;
      content = content.slice(0, at) + span + content.slice(at + token.length);
      placed.add(i);
    });
  }

  const missing = mentions.findIndex((_, i) => !placed.has(i));
  if (missing >= 0) {
    const name = mentions[missing]?.display_name ?? '';
    throw ToolError.validation(
      format === 'html'
        ? `Mention ${String(missing)} ("${name}") is not placed: add {{${String(missing)}}} to the html text`
        : `Mention ${String(missing)} ("${name}") is not placed: include "@${name}" in the text`,
    );
  }

  const mentionsJson = JSON.stringify(
    mentions.map((m, i) => ({
      '@type': MENTION_TYPE,
      itemid: i,
      mri: m.mri,
      mentionType: m.mri.startsWith('19:') ? 'channel' : 'person',
      displayName: m.display_name,
    })),
  );
  return { content, mentionsJson };
};

/** Random numeric client message id (up to 20 digits), as Teams web generates. */
export const randomClientMessageId = (): string => {
  const parts = new Uint32Array(2);
  crypto.getRandomValues(parts);
  const hi = BigInt(parts[0] ?? 0);
  const lo = BigInt(parts[1] ?? 0);
  return ((hi << 32n) | lo).toString();
};

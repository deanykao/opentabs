import { ToolError } from '@opentabs-dev/plugin-sdk';

/**
 * Chat service list endpoints return `_metadata.backwardLink`: an absolute URL
 * for the next (older) page. We hand callers only the path+query from `/v1/`
 * onward so the cursor is host-independent and cannot redirect the Skype token
 * to another origin when it is appended to the chat service base URL.
 */
export const cursorFromBackwardLink = (link: string | undefined): string | undefined => {
  if (!link) return undefined;
  let url: URL;
  try {
    url = new URL(link);
  } catch {
    return undefined;
  }
  const idx = url.pathname.indexOf('/v1/');
  if (idx < 0) return undefined;
  return url.pathname.slice(idx) + url.search;
};

/**
 * Validate a cursor previously returned by `cursorFromBackwardLink` and make
 * sure it carries a page size. `mustContain` pins the cursor to the resource it
 * was issued for (decoded, so a conversation ID matches either encoding).
 */
export const endpointFromCursor = (cursor: string, mustContain: string, pageSize: number): string => {
  let decoded: string;
  try {
    decoded = decodeURIComponent(cursor);
  } catch {
    throw ToolError.validation('Invalid cursor: use the next_cursor value returned by the previous call');
  }
  if (!cursor.startsWith('/v1/users/') || cursor.startsWith('//') || !decoded.includes(mustContain)) {
    throw ToolError.validation('Invalid cursor: use the next_cursor value returned by the previous call');
  }
  if (/[?&]pageSize=/.test(cursor)) return cursor;
  return `${cursor}${cursor.includes('?') ? '&' : '?'}pageSize=${String(pageSize)}`;
};

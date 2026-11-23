import { describe, expect, test } from 'vitest';
import { cursorFromBackwardLink, endpointFromCursor } from './pagination.js';

const link =
  'https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/conversations/19:abc@thread.v2/messages?startTime=0&syncState=XYZ&pageSize=5';

describe('cursorFromBackwardLink', () => {
  test('keeps path and query from /v1/ onward, dropping host and region prefix', () => {
    expect(cursorFromBackwardLink(link)).toBe(
      '/v1/users/ME/conversations/19:abc@thread.v2/messages?startTime=0&syncState=XYZ&pageSize=5',
    );
  });
  test('returns undefined when there is no link', () => {
    expect(cursorFromBackwardLink(undefined)).toBeUndefined();
    expect(cursorFromBackwardLink('')).toBeUndefined();
    expect(cursorFromBackwardLink('https://x.test/other')).toBeUndefined();
  });
});

describe('endpointFromCursor', () => {
  const cursor = cursorFromBackwardLink(link) ?? '';
  test('accepts a cursor for the same conversation', () => {
    expect(endpointFromCursor(cursor, '/conversations/19:abc@thread.v2/messages', 20)).toBe(cursor);
  });
  test('adds pageSize when missing', () => {
    const bare = '/v1/users/ME/conversations?view=msnp24Equivalent&syncState=ABC';
    expect(endpointFromCursor(bare, '/conversations?', 7)).toBe(`${bare}&pageSize=7`);
  });
  test('rejects a cursor issued for another conversation', () => {
    expect(() => endpointFromCursor(cursor, '/conversations/19:other@thread.v2/messages', 20)).toThrow(
      /Invalid cursor/,
    );
  });
  test('rejects absolute and protocol-relative URLs so the token cannot leave the chat service', () => {
    expect(() => endpointFromCursor('https://evil.test/v1/users/ME/conversations?', '/conversations?', 5)).toThrow();
    expect(() => endpointFromCursor('//evil.test/v1/users/ME/conversations?', '/conversations?', 5)).toThrow();
  });
});

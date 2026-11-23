import { describe, expect, test } from 'vitest';
import { mapMessage } from './schemas.js';

describe('mapMessage', () => {
  test('drops reactions whose last user removed them', () => {
    const message = mapMessage({
      id: '1',
      properties: {
        emotions: [
          { key: 'like', users: [] },
          { key: 'heart', users: [{ mri: '8:orgid:x' }] },
        ],
      },
    });
    expect(message.reactions).toEqual(['heart']);
  });

  test('flags deleted messages', () => {
    expect(mapMessage({ id: '1', content: '', properties: { deletetime: 1795461533204 } }).is_deleted).toBe(true);
    expect(mapMessage({ id: '2', content: 'hi' }).is_deleted).toBe(false);
  });
});

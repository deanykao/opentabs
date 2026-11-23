import Ajv from 'ajv';
import { describe, expect, test } from 'vitest';
import { z } from 'zod';
import plugin from './index.js';

/**
 * Mirrors the MCP server's validator (platform/mcp-server/src/registry.ts): strict Ajv
 * without format plugins. A schema it cannot compile makes the tool uncallable, which
 * is how `z.iso.datetime` (emitting `format: "date-time"`) broke every scheduling tool.
 */
const compile = (schema: z.ZodType) => {
  const json = z.toJSONSchema(schema) as Record<string, unknown>;
  delete json.$schema;
  return new Ajv({ allErrors: false }).compile(json);
};

describe('tool schemas compile under the MCP server validator', () => {
  for (const tool of plugin.tools) {
    test(tool.name, () => {
      expect(() => compile(tool.input)).not.toThrow();
      expect(() => compile(tool.output)).not.toThrow();
    });
  }
});

describe('date-time inputs', () => {
  const tool = (name: string) => {
    const found = plugin.tools.find(t => t.name === name);
    if (!found) throw new Error(`missing tool ${name}`);
    return compile(found.input);
  };

  test('local date-time fields accept wall-clock times and reject offsets and bare dates', () => {
    const validate = tool('create_event');
    const base = { subject: 's', end: '2026-06-02T14:00:00' };
    expect(validate({ ...base, start: '2026-06-02T13:00:00' })).toBe(true);
    expect(validate({ ...base, start: '2026-06-02T13:00' })).toBe(true);
    expect(validate({ ...base, start: '2026-06-02T13:00:00-05:00' })).toBe(false);
    expect(validate({ ...base, start: '2026-06-02T13:00:00Z' })).toBe(false);
    expect(validate({ ...base, start: '2026-06-02' })).toBe(false);
    expect(validate({ ...base, start: '2026-13-02T13:00:00' })).toBe(false);
  });

  test('calendar view bounds accept optional offsets', () => {
    const validate = tool('get_calendar_view');
    const base = { end: '2026-06-09T00:00:00-05:00' };
    expect(validate({ ...base, start: '2026-06-02T00:00:00-05:00' })).toBe(true);
    expect(validate({ ...base, start: '2026-06-02T00:00:00Z' })).toBe(true);
    expect(validate({ ...base, start: '2026-06-02T00:00:00' })).toBe(true);
    expect(validate({ ...base, start: '2026-06-02T00:00:00-5' })).toBe(false);
  });

  test('recurrence dates are calendar dates', () => {
    const validate = tool('create_event');
    const event = { subject: 's', start: '2026-06-02T13:00:00', end: '2026-06-02T14:00:00' };
    const recurrence = (start_date: string) => ({
      pattern: { type: 'daily' },
      range: { type: 'noEnd', start_date },
    });
    expect(validate({ ...event, recurrence: recurrence('2026-06-02') })).toBe(true);
    expect(validate({ ...event, recurrence: recurrence('2026-06-02T00:00:00') })).toBe(false);
  });
});

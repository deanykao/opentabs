import Ajv from 'ajv';
import { describe, expect, test } from 'vitest';
import { z } from 'zod';
import plugin from './index.js';

// Mirrors the MCP server's validator (platform/mcp-server/src/registry.ts): an
// unknown `format` keyword (from z.email(), z.url(), z.iso.*, ...) makes compile() throw.
describe('tool schemas compile under Ajv', () => {
  for (const tool of plugin.tools) {
    for (const kind of ['input', 'output'] as const) {
      test(`${tool.name} ${kind}`, () => {
        const schema = z.toJSONSchema(tool[kind]) as Record<string, unknown>;
        delete schema.$schema;
        expect(() => new Ajv({ allErrors: false }).compile(schema)).not.toThrow();
      });
    }
  }
});

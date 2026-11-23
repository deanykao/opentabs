import { z } from 'zod';

/** Shared `format` input for tools that write message bodies. */
export const formatField = z
  .enum(['text', 'html'])
  .optional()
  .describe('"text" (default): plain text, escaped and newlines converted to line breaks. "html": raw HTML');

/** Shared `mentions` input for tools that write message bodies. */
export const mentionsField = z
  .array(
    z.object({
      mri: z.string().min(1).describe('MRI of the mentioned user (e.g., "8:orgid:<guid>")'),
      display_name: z.string().min(1).describe('Name shown for the mention'),
    }),
  )
  .optional()
  .describe(
    'People to @mention. In text format include "@<display_name>" in the text where each mention goes; in html format put {{0}}, {{1}}, ... (zero-based index into this array) where each goes',
  );

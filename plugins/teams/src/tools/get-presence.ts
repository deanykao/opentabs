import { ToolError, defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { getSkypeIdentity, graphApi, presenceApi } from '../teams-api.js';

interface PresenceEntry {
  mri?: string;
  status?: number;
  presence?: {
    availability?: string;
    activity?: string;
    deviceType?: string;
    calendarData?: { isOutOfOffice?: boolean };
    workLocation?: { location?: string };
  };
}

/** Resolve an MRI, email/UPN, or "me" to an MRI. Emails are resolved through Graph (User.ReadBasic.All). */
const resolveMri = async (user: string): Promise<string> => {
  if (user === 'me') return `8:${(await getSkypeIdentity()).skypeid}`;
  if (/^\d+:/.test(user)) return user;
  if (user.includes('@')) {
    const found = await graphApi<{ id?: string }>(`/users/${encodeURIComponent(user)}`, {
      query: { $select: 'id' },
      scope: 'User.ReadBasic.All',
    });
    if (!found.id) throw ToolError.notFound(`No user found for ${user}`);
    return `8:orgid:${found.id}`;
  }
  throw ToolError.validation(`"${user}" is not an MRI (e.g. "8:orgid:<guid>"), an email address, or "me"`);
};

export const getPresence = defineTool({
  name: 'get_presence',
  displayName: 'Get Presence',
  description:
    'Get the Teams presence (Available, Busy, Away, DoNotDisturb, Offline, ...) of one or more users. Each user may be an MRI ("8:orgid:<guid>"), an email address, or "me". Defaults to the current user.',
  summary: 'Get user presence',
  icon: 'circle-dot',
  group: 'People',
  input: z.object({
    users: z
      .array(z.string().min(1))
      .min(1)
      .max(50)
      .optional()
      .describe('MRIs, email addresses, or "me" (default: ["me"])'),
  }),
  output: z.object({
    presence: z.array(
      z.object({
        mri: z.string().describe('User MRI'),
        availability: z.string().describe('Availability (e.g., "Available", "Busy", "Away", "Offline")'),
        activity: z.string().describe('Activity (e.g., "InAMeeting", "Presenting", "Available")'),
        device_type: z.string().describe('Device the presence was reported from (e.g., "Web", "Desktop")'),
        out_of_office: z.boolean().describe('Whether the user is out of office'),
        work_location: z.string().describe('Work location if shared (e.g., "Remote"), else empty'),
        status_code: z.number().describe('Presence service status for this user (20000 = OK)'),
      }),
    ),
  }),
  handle: async params => {
    const mris = await Promise.all((params.users ?? ['me']).map(resolveMri));
    const data = await presenceApi<PresenceEntry[]>('/v1/presence/getpresence/', {
      method: 'POST',
      body: mris.map(mri => ({ mri })),
    });
    return {
      presence: (Array.isArray(data) ? data : []).map(p => ({
        mri: p.mri ?? '',
        availability: p.presence?.availability ?? '',
        activity: p.presence?.activity ?? '',
        device_type: p.presence?.deviceType ?? '',
        out_of_office: p.presence?.calendarData?.isOutOfOffice ?? false,
        work_location: p.presence?.workLocation?.location ?? '',
        status_code: p.status ?? 0,
      })),
    };
  },
});

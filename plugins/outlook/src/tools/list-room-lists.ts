import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { apiPaged } from '../outlook-api.js';
import { emailAddressSchema } from './schemas.js';

/** Graph `place` (displayName/emailAddress) or Outlook REST `EmailAddress` (name/address). */
export interface RawPlace {
  displayName?: string;
  emailAddress?: string;
  name?: string;
  address?: string;
  capacity?: number | null;
  building?: string | null;
  floorNumber?: number | null;
}

export const fetchRoomLists = async () => {
  const { items } = await apiPaged<RawPlace>(
    { graph: '/places/microsoft.graph.roomlist', rest: '/me/findroomlists', restVersion: 'beta' },
    {},
    'calendar',
    1000,
  );
  return items.map(p => ({ name: p.displayName ?? p.name ?? '', address: p.emailAddress ?? p.address ?? '' }));
};

export const listRoomLists = defineTool({
  name: 'list_room_lists',
  displayName: 'List Room Lists',
  description:
    'List the room lists (buildings/locations) defined in your organization. Pass a room list address to list_rooms to see its rooms.',
  summary: 'List room lists',
  icon: 'building',
  group: 'Calendar',
  input: z.object({}),
  output: z.object({
    room_lists: z.array(emailAddressSchema).describe('Room lists (name and mailbox address)'),
  }),
  handle: async () => ({ room_lists: await fetchRoomLists() }),
});

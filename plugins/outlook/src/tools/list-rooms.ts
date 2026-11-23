import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { type ApiEndpoint, apiPaged } from '../outlook-api.js';
import { type RawPlace, fetchRoomLists } from './list-room-lists.js';

const roomSchema = z.object({
  name: z.string().describe('Room display name'),
  address: z.string().describe('Room mailbox address'),
  room_list: z.string().describe('Room list (building/location) the room belongs to'),
  capacity: z.number().nullable().describe('Seating capacity, when published'),
  building: z.string().describe('Building, when published'),
  floor: z.number().nullable().describe('Floor number, when published'),
});

type Room = z.infer<typeof roomSchema>;

/** At most this many room lists whose name matches the query are expanded per call. */
const MAX_MATCHING_LISTS = 5;

const fetchRooms = async (endpoint: ApiEndpoint, roomList: string): Promise<Room[]> => {
  const { items } = await apiPaged<RawPlace>(endpoint, {}, 'calendar', 5000);
  return items.map(p => ({
    name: p.displayName ?? p.name ?? '',
    address: p.emailAddress ?? p.address ?? '',
    room_list: roomList,
    capacity: p.capacity ?? null,
    building: p.building ?? '',
    floor: p.floorNumber ?? null,
  }));
};

// REST beta only accepts the room list through an OData parameter alias; an inline
// literal (`findrooms(RoomList='…')`) fails with "Mailbox info is stale".
const roomListEndpoint = (address: string): ApiEndpoint => ({
  graph: `/places/${encodeURIComponent(address)}/microsoft.graph.roomlist/rooms`,
  rest: `/me/findrooms(RoomList=@list)?@list=${encodeURIComponent(`'${address.replace(/'/g, "''")}'`)}`,
  restVersion: 'beta',
});

export const listRooms = defineTool({
  name: 'list_rooms',
  displayName: 'List Rooms',
  description:
    "List bookable meeting rooms, optionally filtered by text. With room_list (from list_room_lists) every room in that list is returned. Without it, the organization's room directory is searched (Exchange may cap it at 100 rooms) together with up to 5 room lists whose name matches the query — pass room_list for a complete listing. Check availability with get_schedule or find_meeting_times, then book by adding the room as a resource attendee and location_email in create_event.",
  summary: 'List meeting rooms',
  icon: 'door-open',
  group: 'Calendar',
  input: z.object({
    room_list: z.string().optional().describe('Room list mailbox address to restrict to'),
    query: z
      .string()
      .optional()
      .describe('Case-insensitive text the room name, address, or room list name must contain'),
    limit: z.number().int().min(1).max(500).optional().describe('Max rooms to return (default 100)'),
  }),
  output: z.object({
    rooms: z.array(roomSchema).describe('Matching rooms'),
    has_more: z.boolean().describe('True when more rooms matched than limit'),
  }),
  handle: async params => {
    const limit = params.limit ?? 100;
    const needle = params.query?.toLowerCase();

    // Sequential on purpose: fanning out over every room list trips Exchange throttling.
    const rooms: Room[] = [];
    if (params.room_list) {
      rooms.push(...(await fetchRooms(roomListEndpoint(params.room_list), params.room_list)));
    } else {
      rooms.push(
        ...(await fetchRooms(
          { graph: '/places/microsoft.graph.room', rest: '/me/findrooms', restVersion: 'beta' },
          '',
        )),
      );
      if (needle) {
        const lists = (await fetchRoomLists())
          .filter(l => l.name.toLowerCase().includes(needle) || l.address.toLowerCase().includes(needle))
          .slice(0, MAX_MATCHING_LISTS);
        for (const list of lists) rooms.push(...(await fetchRooms(roomListEndpoint(list.address), list.name)));
      }
    }

    // Neither API filters room names server-side. Rooms seen in a list replace the
    // directory entry so room_list is filled in.
    const byAddress = new Map<string, Room>();
    for (const room of rooms) {
      const key = room.address.toLowerCase();
      if (!byAddress.has(key) || room.room_list) byAddress.set(key, room);
    }
    const matches = [...byAddress.values()].filter(
      r => !needle || [r.name, r.address, r.room_list].some(s => s.toLowerCase().includes(needle)),
    );
    return { rooms: matches.slice(0, limit), has_more: matches.length > limit };
  },
});

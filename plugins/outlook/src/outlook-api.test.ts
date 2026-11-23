import { describe, expect, test } from 'vitest';
import { pascalCaseKeys, rankCandidates, scopesForHost } from './outlook-api.js';

describe('scopesForHost', () => {
  test('returns lowercased scope names for the host, stripping leading slashes', () => {
    const target =
      'https://outlook.office.com//Calendars.ReadWrite https://outlook.office.com/Mail.Send openid https://graph.microsoft.com/User.Read';
    expect(scopesForHost(target, 'outlook.office.com')).toEqual(['calendars.readwrite', 'mail.send']);
  });

  test('ignores scopes that only mention the host in their path', () => {
    expect(scopesForHost('https://attacker.com/graph.microsoft.com/Calendars.Read', 'graph.microsoft.com')).toEqual([]);
  });
});

describe('rankCandidates', () => {
  // The shape observed in an enterprise Outlook tab: a User-only Graph token and a
  // broad Outlook REST token. Calendar calls must go to the REST token.
  const graphUserOnly = { id: 'graph', scopes: ['user.read', 'user.readbasic.all'] };
  const restCalendar = { id: 'rest', scopes: ['calendars.read', 'calendars.readwrite', 'mail.readwrite'] };
  const restReadOnly = { id: 'rest-ro', scopes: ['calendars.read'] };
  const defaultScope = { id: 'default', scopes: ['.default'] };
  const v1 = { id: 'v1', scopes: undefined };

  test('calendar reads skip tokens without calendar scopes and try granting tokens first', () => {
    const ids = rankCandidates([graphUserOnly, defaultScope, v1, restCalendar], 'calendar').map(c => c.id);
    expect(ids).toEqual(['rest', 'default', 'v1']);
  });

  test('calendar writes require a ReadWrite scope', () => {
    const ids = rankCandidates([restReadOnly, restCalendar], 'calendar-write').map(c => c.id);
    expect(ids).toEqual(['rest']);
  });

  test('mail keeps every candidate in discovery order', () => {
    const ids = rankCandidates([graphUserOnly, restCalendar, v1], 'mail').map(c => c.id);
    expect(ids).toEqual(['graph', 'rest', 'v1']);
  });
});

describe('pascalCaseKeys', () => {
  test('uses REST v2.0 names where they differ from PascalCase', () => {
    expect(
      pascalCaseKeys({ timeConstraint: { activityDomain: 'work', timeSlots: [{ start: { dateTime: 'x' } }] } }),
    ).toEqual({ TimeConstraint: { ActivityDomain: 'work', Timeslots: [{ Start: { DateTime: 'x' } }] } });
  });
});

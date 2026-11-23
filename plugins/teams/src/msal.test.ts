import { describe, expect, test } from 'vitest';
import { scopeClaimHasHost, selectTokens } from './msal.js';

const NOW = 1_700_000_000_000;
const exp = (secondsFromNow: number): string => String(NOW / 1000 + secondsFromNow);

describe('scopeClaimHasHost', () => {
  test('matches by hostname only', () => {
    expect(scopeClaimHasHost('openid https://graph.microsoft.com/User.Read', 'graph.microsoft.com')).toBe(true);
    expect(scopeClaimHasHost('https://attacker.com/graph.microsoft.com/foo', 'graph.microsoft.com')).toBe(false);
    expect(scopeClaimHasHost('openid profile', 'graph.microsoft.com')).toBe(false);
  });
});

describe('selectTokens', () => {
  const g = (name: string) => `https://graph.microsoft.com/${name}`;

  test('drops expired, headroom-less, empty and wrong-host entries; sorts longest-lived first', () => {
    const tokens = selectTokens(
      [
        { secret: 'expired', target: g('A'), expiresOn: exp(-5) },
        { secret: 'tight', target: g('A'), expiresOn: exp(20) },
        { secret: 'short', target: g('A'), expiresOn: exp(600) },
        { secret: 'long', target: g('A'), expiresOn: exp(3000) },
        { secret: '', target: g('A'), expiresOn: exp(3000) },
        { secret: 'other-host', target: 'https://outlook.office.com/x', expiresOn: exp(3000) },
        { secret: 'junk-expiry', target: g('A'), expiresOn: '9999999999junk' },
      ],
      'graph.microsoft.com',
      NOW,
      60,
    );
    expect(tokens.map(t => t.secret)).toEqual(['long', 'short']);
  });

  test('a token carrying the preferred scope wins over a longer-lived one without it', () => {
    const tokens = selectTokens(
      [
        { secret: 'calendars', target: g('Calendars.ReadWrite'), expiresOn: exp(3000) },
        {
          secret: 'channels',
          target: `${g('Team.ReadBasic.All')} ${g('ChannelMessage.Read.All')}`,
          expiresOn: exp(1000),
        },
      ],
      'graph.microsoft.com',
      NOW,
      0,
      'ChannelMessage.Read.All',
    );
    expect(tokens.map(t => t.secret)).toEqual(['channels', 'calendars']);
  });
});

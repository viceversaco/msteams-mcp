import { beforeEach, describe, expect, it, vi } from 'vitest';
import { extractSubstrateToken, getValidSubstrateToken } from './token-extractor.js';
import { readSessionState, readTokenCache, writeTokenCache, type SessionState } from './session-store.js';
import { SUBSTRATE_AUDIENCE } from '../constants.js';

vi.mock('./session-store.js', async importOriginal => ({
  ...await importOriginal<typeof import('./session-store.js')>(),
  readSessionState: vi.fn(), readTokenCache: vi.fn(), writeTokenCache: vi.fn(),
}));

const OUTLOOK_SEARCH = 'https://outlook.office.com/search/SubstrateSearch-Internal.ReadWrite';
const SUBSTRATE_SEARCH = 'https://substrate.office.com/SubstrateSearch-Internal.ReadWrite';
const CLIENT_AUD = '00000000-0000-0000-0000-0000000000c1';

function token(aud: string, seconds = 3600) {
  return `eyJhbGciOiJub25lIn0.${Buffer.from(JSON.stringify({ aud, exp: Date.now() / 1000 + seconds })).toString('base64url')}.signature`;
}
function entry(secret: string, target: string) {
  return { name: secret, value: JSON.stringify({ credentialType: 'AccessToken', target, secret }) };
}
function state(entries: ReturnType<typeof entry>[]): SessionState {
  return { cookies: [], origins: [{ origin: 'https://teams.microsoft.com', localStorage: entries }] };
}

describe('Substrate token selection', () => {
  it('prefers the substrate-audience token over a longer-lived outlook-audience one', () => {
    const substrate = token(SUBSTRATE_AUDIENCE, 3600);
    const outlook = token(CLIENT_AUD, 5000);
    const picked = extractSubstrateToken(state([entry(outlook, OUTLOOK_SEARCH), entry(substrate, SUBSTRATE_SEARCH)]));
    expect(picked?.token).toBe(substrate);
  });
  it('falls back to the longest-lived SubstrateSearch token when none has the substrate audience', () => {
    const shorter = token(CLIENT_AUD, 1000);
    const longer = token(CLIENT_AUD, 5000);
    expect(extractSubstrateToken(state([entry(shorter, OUTLOOK_SEARCH), entry(longer, OUTLOOK_SEARCH)]))?.token).toBe(longer);
  });
  it('ignores an expired substrate-audience token', () => {
    const outlook = token(CLIENT_AUD, 5000);
    const picked = extractSubstrateToken(state([entry(token(SUBSTRATE_AUDIENCE, -1), SUBSTRATE_SEARCH), entry(outlook, OUTLOOK_SEARCH)]));
    expect(picked?.token).toBe(outlook);
  });
  it('ignores a substrate-audience token whose target is not SubstrateSearch', () => {
    const other = token(SUBSTRATE_AUDIENCE, 5000);
    const search = token(CLIENT_AUD, 1000);
    expect(extractSubstrateToken(state([entry(other, 'https://substrate.office.com/.default'), entry(search, OUTLOOK_SEARCH)]))?.token).toBe(search);
  });
});

describe('cached Substrate token', () => {
  beforeEach(() => vi.clearAllMocks());
  const cacheOf = (secret: string) => ({ substrateToken: secret, substrateTokenExpiry: Date.now() + 3.6e6, extractedAt: Date.now() });

  it('replaces a cached outlook-audience token once the session has a substrate-audience one', () => {
    const outlook = token(CLIENT_AUD, 5000);
    const substrate = token(SUBSTRATE_AUDIENCE, 3600);
    vi.mocked(readTokenCache).mockReturnValue(cacheOf(outlook));
    vi.mocked(readSessionState).mockReturnValue(state([entry(outlook, OUTLOOK_SEARCH), entry(substrate, SUBSTRATE_SEARCH)]));
    expect(getValidSubstrateToken()).toBe(substrate);
    expect(vi.mocked(writeTokenCache).mock.calls[0][0].substrateToken).toBe(substrate);
  });

  it('keeps serving a cached substrate-audience token without reading the session', () => {
    const substrate = token(SUBSTRATE_AUDIENCE, 3600);
    vi.mocked(readTokenCache).mockReturnValue(cacheOf(substrate));
    expect(getValidSubstrateToken()).toBe(substrate);
    expect(readSessionState).not.toHaveBeenCalled();
  });

  it('keeps the cached token when the session has nothing better', () => {
    const outlook = token(CLIENT_AUD, 5000);
    vi.mocked(readTokenCache).mockReturnValue(cacheOf(outlook));
    vi.mocked(readSessionState).mockReturnValue(state([entry(outlook, OUTLOOK_SEARCH)]));
    expect(getValidSubstrateToken()).toBe(outlook);
    expect(writeTokenCache).not.toHaveBeenCalled();
  });
});

import { beforeEach, describe, expect, it, vi } from 'vitest';
import { getSharedFiles, type SharedFile } from './files-api.js';
import { httpRequest } from '../utils/http.js';
import { requireGraphTokenAsync } from '../utils/auth-guards.js';
import { ok } from '../types/result.js';
import { findMeetingRecording, getRecordingInfo } from './recording-api.js';

vi.mock('./files-api.js', () => ({ getSharedFiles: vi.fn() }));
vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ requireGraphTokenAsync: vi.fn() }));

const SP = 'https://contoso-my.sharepoint.com/personal/owner/Documents/Recordings';

function mp4(name: string, sharedTime: string, sharedBy = 'Owner'): SharedFile {
  return { itemType: 'File', fileName: name, webUrl: `${SP}/${encodeURIComponent(name)}`, sharedTime, sharedBy };
}
function pages(...lists: SharedFile[][]) {
  const mock = vi.mocked(getSharedFiles);
  lists.forEach((files, i) => mock.mockResolvedValueOnce(ok({ files, skipToken: i < lists.length - 1 ? `p${i + 1}` : undefined } as never)));
}

beforeEach(() => {
  vi.clearAllMocks();
  vi.mocked(requireGraphTokenAsync).mockResolvedValue(ok('graph-token'));
});

describe('findMeetingRecording', () => {
  const early = mp4('Sync-20260105_160000-Meeting Recording.mp4', '2026-01-05T16:50:00Z');
  const late = mp4('Sync-20260112_160000-Meeting Recording.mp4', '2026-01-12T16:50:00Z');
  const clip = mp4('demo.mp4', '2026-01-13T09:00:00Z');

  it('pages through the listing and picks the newest meeting recording', async () => {
    pages([early, clip], [late]);
    const result = await findMeetingRecording('19:meeting_x@thread.v2');
    expect(vi.mocked(getSharedFiles)).toHaveBeenLastCalledWith('19:meeting_x@thread.v2', { pageSize: 100, skipToken: 'p1' });
    expect(result.ok && result.value.recording?.url).toBe(late.webUrl);
    expect(result.ok && result.value.recording?.candidates).toHaveLength(2);
  });

  it('picks the recording nearest to the given date', async () => {
    pages([early, late]);
    const result = await findMeetingRecording('19:meeting_x@thread.v2', '2026-01-05T16:00:00Z');
    expect(result.ok && result.value.recording?.url).toBe(early.webUrl);
  });

  it('falls back to any .mp4 when none is named Meeting Recording', async () => {
    pages([clip, { itemType: 'Link', webUrl: 'https://example.com', title: 'x' }]);
    const result = await findMeetingRecording('19:meeting_x@thread.v2');
    expect(result.ok && result.value.recording?.url).toBe(clip.webUrl);
  });

  it('stops when the listing repeats a page token', async () => {
    vi.mocked(getSharedFiles).mockResolvedValue(ok({ files: [early], skipToken: 'again' } as never));
    const result = await findMeetingRecording('19:meeting_x@thread.v2');
    expect(result.ok).toBe(false);
    expect(vi.mocked(getSharedFiles)).toHaveBeenCalledTimes(2);
  });

  it('stops after a bounded number of pages', async () => {
    let n = 0;
    vi.mocked(getSharedFiles).mockImplementation(async () => ok({ files: [], skipToken: `p${++n}` } as never));
    const result = await findMeetingRecording('19:meeting_x@thread.v2');
    expect(result.ok).toBe(false);
    expect(n).toBe(50);
  });

  it('returns null with the shared names when there is no .mp4', async () => {
    pages([{ itemType: 'File', fileName: 'notes.docx', webUrl: `${SP}/notes.docx` }]);
    const result = await findMeetingRecording('19:meeting_x@thread.v2');
    expect(result.ok && result.value).toEqual({ recording: null, files: ['notes.docx'] });
  });
});

describe('getRecordingInfo', () => {
  it('maps Graph metadata and the blocksDownload flag', async () => {
    vi.mocked(httpRequest).mockResolvedValue(ok({ status: 200, headers: new Headers(), data: {
      name: 'r.mp4', size: 10, createdDateTime: '2026-01-12T17:00:00Z',
      file: { hashes: { quickXorHash: 'qx' } }, video: { duration: 3600000, width: 1920, height: 1080 },
      currentUserRole: { blocksDownload: true },
    } }));
    const url = `${SP}/r.mp4`;
    const result = await getRecordingInfo(url, { url, sharedBy: 'Owner', candidates: [{ webUrl: url }] });
    expect(result.ok && result.value).toEqual({
      name: 'r.mp4', webUrl: url, size: 10, quickXorHash: 'qx', durationMs: 3600000, width: 1920, height: 1080,
      sharedBy: 'Owner', created: '2026-01-12T17:00:00Z', blocksDownload: true, candidates: undefined,
    });
    expect(vi.mocked(httpRequest).mock.calls[0][0]).toMatch(/^https:\/\/graph\.microsoft\.com\/v1\.0\/shares\/u!/);
  });

  it('reports blocksDownload null when Graph omits currentUserRole', async () => {
    vi.mocked(httpRequest).mockResolvedValue(ok({ status: 200, headers: new Headers(), data: { name: 'r.mp4', size: 10 } }));
    const result = await getRecordingInfo(`${SP}/r.mp4`);
    expect(result.ok && result.value.blocksDownload).toBeNull();
  });

  it('never sends the Graph token to a non-SharePoint URL', async () => {
    const result = await getRecordingInfo('https://evil.example.com/r.mp4');
    expect(result.ok).toBe(false);
    expect(vi.mocked(httpRequest)).not.toHaveBeenCalled();
    expect(vi.mocked(requireGraphTokenAsync)).not.toHaveBeenCalled();
  });
});

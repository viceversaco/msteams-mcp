import { beforeEach, describe, expect, it, vi } from 'vitest';
import { getTranscriptContent } from '../api/transcript-api.js';
import { findMeetingRecording, getRecordingInfo, type RecordingInfo } from '../api/recording-api.js';
import { downloadFile } from '../api/graph-files-api.js';
import { ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { runMeetingCommand } from './meeting.js';

vi.mock('../api/transcript-api.js', () => ({ getTranscriptContent: vi.fn() }));
vi.mock('../api/recording-api.js', () => ({ findMeetingRecording: vi.fn(), getRecordingInfo: vi.fn() }));
vi.mock('../api/graph-files-api.js', () => ({ downloadFile: vi.fn() }));

const THREAD = '19:meeting_x@thread.v2';
const URL_ = 'https://contoso-my.sharepoint.com/personal/o/Documents/Recordings/r.mp4';

function info(blocksDownload: boolean | null): RecordingInfo {
  return { name: 'r.mp4', webUrl: URL_, size: 10, blocksDownload };
}
function located() {
  vi.mocked(findMeetingRecording).mockResolvedValue(ok({ recording: { url: URL_, sharedBy: 'Owner', candidates: [] }, files: ['r.mp4'] }));
}

beforeEach(() => vi.clearAllMocks());

describe('transcript', () => {
  it('prints the transcript without formattedText and passes the date through', async () => {
    vi.mocked(getTranscriptContent).mockResolvedValue(ok({
      threadId: THREAD, entries: [{ startTime: '00:00:01', endTime: '00:00:02', speaker: 'A', text: 'hi' }],
      formattedText: 'A: hi', entryCount: 1, speakers: ['A'],
    } as never));
    const { code, out } = await runMeetingCommand(['transcript', THREAD, '2026-01-12T16:00:00Z']);
    expect(code).toBe(0);
    expect(getTranscriptContent).toHaveBeenCalledWith(THREAD, '2026-01-12T16:00:00Z');
    expect(out).toEqual({ ok: true, threadId: THREAD, entries: [{ startTime: '00:00:01', endTime: '00:00:02', speaker: 'A', text: 'hi' }], entryCount: 1, speakers: ['A'] });
  });

  it('exits 1 with the API error', async () => {
    vi.mocked(getTranscriptContent).mockResolvedValue(err(createError(ErrorCode.NOT_FOUND, 'No recording found')));
    expect(await runMeetingCommand(['transcript', THREAD])).toEqual({ code: 1, out: { ok: false, error: 'No recording found', code: 'NOT_FOUND' } });
  });
});

describe('recording', () => {
  it('prints metadata with --info and never downloads', async () => {
    located();
    vi.mocked(getRecordingInfo).mockResolvedValue(ok(info(false)));
    const { code, out } = await runMeetingCommand(['recording', THREAD, '--info']);
    expect(code).toBe(0);
    expect(out).toMatchObject({ ok: true, name: 'r.mp4', blocksDownload: false });
    expect(downloadFile).not.toHaveBeenCalled();
  });

  it('exits 4 and writes nothing when the owner blocks download', async () => {
    located();
    vi.mocked(getRecordingInfo).mockResolvedValue(ok(info(true)));
    const { code } = await runMeetingCommand(['recording', THREAD, '--out', '/tmp/x']);
    expect(code).toBe(4);
    expect(downloadFile).not.toHaveBeenCalled();
  });

  it('exits 3 when the chat shares no .mp4', async () => {
    vi.mocked(findMeetingRecording).mockResolvedValue(ok({ recording: null, files: ['notes.docx'] }));
    expect(await runMeetingCommand(['recording', THREAD, '--info'])).toEqual({
      code: 3, out: { ok: false, error: 'no .mp4 recording is shared in this meeting chat', files: ['notes.docx'] },
    });
  });

  it('downloads an allowed recording by URL into --out', async () => {
    vi.mocked(getRecordingInfo).mockResolvedValue(ok(info(false)));
    vi.mocked(downloadFile).mockResolvedValue(ok({ name: 'r.mp4', outputPath: '/tmp/x/r.mp4', size: 10, contentType: 'video/mp4', sha256: 'h' }));
    const { code, out } = await runMeetingCommand(['recording', URL_, '--out', '/tmp/x']);
    expect(code).toBe(0);
    expect(findMeetingRecording).not.toHaveBeenCalled();
    expect(vi.mocked(downloadFile).mock.calls[0]).toEqual([URL_, expect.stringMatching(/[\\/]tmp[\\/]x[\\/]r\.mp4$/)]);
    expect(out).toMatchObject({ ok: true, outputPath: '/tmp/x/r.mp4', sha256: 'h' });
  });

  it('requires --out to download, before any network call', async () => {
    expect((await runMeetingCommand(['recording', URL_])).code).toBe(2);
    expect(getRecordingInfo).not.toHaveBeenCalled();
  });

  it('exits 4 and writes nothing when Graph does not report the download right', async () => {
    vi.mocked(getRecordingInfo).mockResolvedValue(ok({ ...info(false), blocksDownload: null }));
    const { code } = await runMeetingCommand(['recording', URL_, '--out', '/tmp/x']);
    expect(code).toBe(4);
    expect(downloadFile).not.toHaveBeenCalled();
  });
});

describe('usage', () => {
  it.each([
    [[]],
    [['transcript']],
    [['bogus', THREAD]],
    [['transcript', THREAD, 'date', 'extra']],
    [['recording', THREAD, '--date']],
    [['recording', THREAD, '--force']],
    [['transcript', THREAD, 'not-a-date']],
    [['recording', THREAD, '--info', '--date', 'soon']],
    [['recording', THREAD, '--out', '--info']],
    [['recording', THREAD, '--out', 'relative/dir']],
  ])('exits 2 for %j before any network call', async argv => {
    expect((await runMeetingCommand(argv)).code).toBe(2);
    expect(getTranscriptContent).not.toHaveBeenCalled();
    expect(findMeetingRecording).not.toHaveBeenCalled();
    expect(getRecordingInfo).not.toHaveBeenCalled();
  });
});

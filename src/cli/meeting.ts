/**
 * Meeting transcript and recording commands for meeting-cli.
 *
 *   transcript <threadId> [meetingDate]
 *   recording <threadId|url> [--date ISO] [--info] [--out DIR]
 *
 * Output is one JSON object: `{ ok: true, ... }` or `{ ok: false, error, ... }`.
 * transcript is getTranscriptContent's value without formattedText. recording
 * prints Graph metadata and, with --out, downloads into DIR unless the owner
 * blocks it. Exit codes: 0 ok, 1 API error, 2 usage, 3 no .mp4 in the meeting
 * chat, 4 the owner blocks download (nothing written).
 */

import { isAbsolute, join } from 'node:path';
import { getTranscriptContent } from '../api/transcript-api.js';
import { findMeetingRecording, getRecordingInfo, type LocatedRecording } from '../api/recording-api.js';
import { downloadFile } from '../api/graph-files-api.js';

export const USAGE = 'usage: transcript <threadId> [date] | recording <threadId|url> [--date ISO] [--info] [--out DIR]';

export interface CommandOutput {
  code: number;
  out: Record<string, unknown>;
}

class Fail extends Error {
  constructor(readonly code: number, message: string, readonly extra: Record<string, unknown> = {}) {
    super(message);
  }
}

async function transcript(threadId: string, meetingDate?: string): Promise<Record<string, unknown>> {
  const result = await getTranscriptContent(threadId, meetingDate);
  if (!result.ok) throw new Fail(1, result.error.message, { code: result.error.code });
  const { formattedText: _formattedText, ...value } = result.value;
  return value;
}

async function recording(target: string, opts: { date?: string; info: boolean; out?: string }): Promise<Record<string, unknown>> {
  let located: LocatedRecording | undefined;
  if (!/^https:\/\//.test(target)) {
    const search = await findMeetingRecording(target, opts.date);
    if (!search.ok) throw new Fail(1, search.error.message, { code: search.error.code });
    if (!search.value.recording) {
      throw new Fail(3, 'no .mp4 recording is shared in this meeting chat', { files: search.value.files });
    }
    located = search.value.recording;
  }
  const url = located?.url ?? target;

  const info = await getRecordingInfo(url, located);
  if (!info.ok) throw new Fail(1, info.error.message, { code: info.error.code, url });
  if (opts.info) return { ...info.value };
  if (info.value.blocksDownload) {
    throw new Fail(4, 'The owner has blocked downloading this recording for you (view-only). Ask the owner to allow download, or download it by hand in the browser.', { ...info.value });
  }
  if (!info.value.name) throw new Fail(1, 'Graph returned no file name for this recording', { ...info.value });

  const saved = await downloadFile(url, join(opts.out!, info.value.name));
  if (!saved.ok) throw new Fail(1, saved.error.message, { code: saved.error.code, ...info.value });
  return { ...info.value, ...saved.value };
}

async function dispatch(argv: string[]): Promise<Record<string, unknown>> {
  const [command, target, ...rest] = argv;
  const opts: { date?: string; info: boolean; out?: string } = { info: false };
  for (let i = 0; i < rest.length; i++) {
    const arg = rest[i];
    if (command === 'recording' && (arg === '--date' || arg === '--out')) {
      const value = rest[++i];
      if (value === undefined) throw new Fail(2, `${arg} needs a value`);
      if (arg === '--date') opts.date = value; else opts.out = value;
    } else if (command === 'recording' && arg === '--info') {
      opts.info = true;
    } else if (command === 'transcript' && i === 0 && !arg.startsWith('-')) {
      opts.date = arg;
    } else {
      throw new Fail(2, `unknown argument: ${arg}`);
    }
  }
  if (!target || target.startsWith('-')) throw new Fail(2, USAGE);
  if (opts.date !== undefined && Number.isNaN(Date.parse(opts.date))) throw new Fail(2, `not a date: ${opts.date}`);
  if (command === 'recording' && !opts.info) {
    if (!opts.out) throw new Fail(2, 'recording needs --info, or --out <absolute directory> to download');
    if (opts.out.startsWith('-') || !isAbsolute(opts.out)) throw new Fail(2, `--out must be an absolute directory: ${opts.out}`);
  }
  if (command === 'transcript') return transcript(target, opts.date);
  if (command === 'recording') return recording(target, opts);
  throw new Fail(2, `unknown command: ${command ?? ''}. ${USAGE}`);
}

export async function runMeetingCommand(argv: string[]): Promise<CommandOutput> {
  try {
    return { code: 0, out: { ok: true, ...(await dispatch(argv)) } };
  } catch (error) {
    const fail = error instanceof Fail ? error : new Fail(1, error instanceof Error ? error.message : String(error));
    return { code: fail.code, out: { ok: false, error: fail.message, ...fail.extra } };
  }
}

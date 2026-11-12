/**
 * Meeting recordings: find the .mp4 a meeting chat shares and read its Graph
 * metadata, including whether the owner lets this user download it.
 *
 * Downloading goes through downloadFile(), which refuses an item whose owner
 * blocks downloads. Recordings created by Teams report createdBy
 * "SharePoint App", so the sharer comes from the chat's file listing.
 */

import { getSharedFiles, type SharedFile } from './files-api.js';
import { toDriveItemUrl } from './graph-files-api.js';
import { httpRequest } from '../utils/http.js';
import { requireGraphTokenAsync } from '../utils/auth-guards.js';
import { type Result, ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';

/** Listing pages read before giving up on a chat whose paging never ends. */
const MAX_LISTING_PAGES = 50;

export interface RecordingCandidate {
  fileName?: string;
  sharedTime?: string;
  webUrl?: string;
}

export interface LocatedRecording {
  url: string;
  sharedBy?: string;
  candidates: RecordingCandidate[];
}

export interface RecordingSearch {
  /** The chosen recording, or null when the chat shares no .mp4. */
  recording: LocatedRecording | null;
  /** Name (or URL) of every item the chat shares. */
  files: string[];
}

export interface RecordingInfo {
  name?: string;
  webUrl: string;
  size?: number;
  quickXorHash?: string;
  durationMs?: number;
  width?: number;
  height?: number;
  sharedBy?: string;
  created?: string;
  /** null when Graph does not say; treat that as blocked. */
  blocksDownload: boolean | null;
  /** Present when the chat shares more than one candidate. */
  candidates?: RecordingCandidate[];
}

interface RawRecordingItem {
  name?: string;
  size?: number;
  createdDateTime?: string;
  file?: { hashes?: { quickXorHash?: string } };
  video?: { duration?: number; width?: number; height?: number };
  currentUserRole?: { blocksDownload?: boolean };
}

/**
 * Picks the recording a meeting chat shares: .mp4 files named "Meeting
 * Recording" first (any .mp4 otherwise), nearest to `date` when given, else
 * the newest.
 */
export async function findMeetingRecording(threadId: string, date?: string): Promise<Result<RecordingSearch>> {
  const files: SharedFile[] = [];
  const seen = new Set<string>();
  let skipToken: string | undefined;
  do {
    if (seen.size >= MAX_LISTING_PAGES) {
      return err(createError(ErrorCode.API_ERROR, `The chat's file listing did not end after ${MAX_LISTING_PAGES} pages`, { retryable: false }));
    }
    const page = await getSharedFiles(threadId, { pageSize: 100, skipToken });
    if (!page.ok) return page;
    files.push(...page.value.files);
    skipToken = page.value.skipToken;
    if (skipToken && seen.has(skipToken)) {
      return err(createError(ErrorCode.API_ERROR, "The chat's file listing repeated a page token", { retryable: false }));
    }
    if (skipToken) seen.add(skipToken);
  } while (skipToken);

  const videos = files.filter(f => f.itemType === 'File' && /\.mp4$/i.test(f.fileName ?? '') && f.webUrl);
  const named = videos.filter(f => /Meeting Recording/i.test(f.fileName ?? ''));
  const pool = (named.length ? named : videos).map(f => ({ file: f, t: Date.parse(f.sharedTime ?? '') || 0 }));
  const names = files.map(f => f.fileName ?? f.webUrl ?? '');
  if (pool.length === 0) return ok({ recording: null, files: names });

  const when = date ? Date.parse(date) : NaN;
  pool.sort((a, b) => Number.isNaN(when) ? b.t - a.t : Math.abs(a.t - when) - Math.abs(b.t - when));
  return ok({
    recording: {
      url: pool[0].file.webUrl!,
      sharedBy: pool[0].file.sharedBy,
      candidates: pool.map(({ file }) => ({ fileName: file.fileName, sharedTime: file.sharedTime, webUrl: file.webUrl })),
    },
    files: names,
  });
}

/** Graph metadata for a recording URL (SharePoint/OneDrive web URL or Graph drive item). */
export async function getRecordingInfo(url: string, located?: LocatedRecording): Promise<Result<RecordingInfo>> {
  const itemUrl = toDriveItemUrl(url);
  if (!itemUrl) {
    return err(createError(ErrorCode.INVALID_INPUT, 'url must be an HTTPS SharePoint/OneDrive file URL or a Graph drive-item URL'));
  }
  const token = await requireGraphTokenAsync();
  if (!token.ok) return token;

  const params = new URLSearchParams({ '$select': 'id,name,size,file,video,createdDateTime,currentUserRole' });
  const meta = await httpRequest<RawRecordingItem>(`${itemUrl}?${params.toString()}`, {
    method: 'GET',
    headers: { Authorization: `Bearer ${token.value}`, Accept: 'application/json' },
  });
  if (!meta.ok) return meta;

  const item = meta.value.data;
  const candidates = located?.candidates ?? [];
  return ok({
    name: item.name,
    webUrl: url,
    size: item.size,
    quickXorHash: item.file?.hashes?.quickXorHash,
    durationMs: item.video?.duration,
    width: item.video?.width,
    height: item.video?.height,
    sharedBy: located?.sharedBy,
    created: item.createdDateTime,
    blocksDownload: item.currentUserRole?.blocksDownload ?? null,
    candidates: candidates.length > 1 ? candidates : undefined,
  });
}

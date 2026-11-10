/**
 * Read-only mode, the default for this build.
 *
 * Only tools in READ_TOOLS are listed or callable. MSTEAMS_ALLOW_WRITES=1
 * restores upstream behaviour: every tool, including unclassified ones.
 * Parameters of read tools that write are stripped from the listed schema and
 * refused on call while writes are off.
 *
 * read-only.test.ts fails when a registered tool is in neither set, so a tool
 * added upstream stays hidden and blocks the build until it is classified.
 */

import type { Tool } from '@modelcontextprotocol/sdk/types.js';

export const READ_TOOLS: ReadonlySet<string> = new Set([
  'teams_login',
  'teams_status',
  'teams_get_me',
  'teams_search_people',
  'teams_get_frequent_contacts',
  'teams_get_person',
  'teams_search',
  'teams_get_thread',
  'teams_find_channel',
  'teams_search_email',
  'teams_get_favorites',
  'teams_get_chat',
  'teams_get_unread',
  'teams_get_activity',
  'teams_search_emoji',
  'teams_get_saved_messages',
  'teams_get_followed_threads',
  'teams_get_message',
  'teams_list_chats',
  'teams_wait_for_reply',
  'teams_get_meetings',
  'teams_get_transcript',
  'teams_get_shared_files',
  'teams_get_tags',
  'teams_list_assignments',
  'teams_get_assignment',
]);

export const WRITE_TOOLS: ReadonlySet<string> = new Set([
  'teams_send_message',
  'teams_edit_message',
  'teams_delete_message',
  'teams_add_reaction',
  'teams_remove_reaction',
  'teams_mark_read',
  'teams_add_favorite',
  'teams_remove_favorite',
  'teams_save_message',
  'teams_unsave_message',
  'teams_create_group_chat',
  'teams_submission_action',
  'teams_upload_file',
  'teams_download_file',
]);

const WRITE_PARAMS: Record<string, string[]> = {
  teams_get_thread: ['markRead'],
};

export function writesEnabled(): boolean {
  return process.env.MSTEAMS_ALLOW_WRITES === '1';
}

export function isToolAllowed(name: string): boolean {
  return writesEnabled() || READ_TOOLS.has(name);
}

/** The definition as listed: write parameters removed while writes are off. */
export function exposedDefinition(definition: Tool): Tool {
  const params = WRITE_PARAMS[definition.name];
  if (!params || writesEnabled()) return definition;
  const properties = { ...(definition.inputSchema.properties ?? {}) };
  for (const param of params) delete properties[param];
  return { ...definition, inputSchema: { ...definition.inputSchema, properties } };
}

/** Why a call is refused, or null when it may run. */
export function refusalReason(name: string, args: unknown): string | null {
  if (writesEnabled()) return null;
  if (!READ_TOOLS.has(name)) {
    return `${name} is disabled: this build is read-only (MSTEAMS_ALLOW_WRITES=1 enables write tools).`;
  }
  const values = (args ?? {}) as Record<string, unknown>;
  const used = (WRITE_PARAMS[name] ?? []).find(param => values[param]);
  return used ? `${name} ${used} is disabled: this build is read-only (MSTEAMS_ALLOW_WRITES=1 enables it).` : null;
}

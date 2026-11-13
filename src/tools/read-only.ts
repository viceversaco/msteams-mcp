/**
 * Read-only mode, the default for this build.
 *
 * Only tools in READ_TOOLS are listed or callable. MSTEAMS_ALLOW_WRITES=1
 * restores upstream behaviour: every tool, including unclassified ones.
 * Parameters of read tools that write are stripped from the listed schema and
 * refused on call until their variable is set. teams_login's forceNew deletes
 * the sign-in every client on the machine shares, so it has its own variable
 * for the one legitimate use, switching accounts.
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

const GUARDED_PARAMS: Record<string, Array<{ param: string; env: string }>> = {
  teams_get_thread: [{ param: 'markRead', env: 'MSTEAMS_ALLOW_WRITES' }],
  teams_login: [{ param: 'forceNew', env: 'MSTEAMS_ALLOW_FORCE_LOGIN' }],
};

export function writesEnabled(): boolean {
  return process.env.MSTEAMS_ALLOW_WRITES === '1';
}

const unlocked = (env: string) => writesEnabled() || process.env[env] === '1';

export function isToolAllowed(name: string): boolean {
  return writesEnabled() || READ_TOOLS.has(name);
}

/** The definition as listed: guarded parameters removed while locked. */
export function exposedDefinition(definition: Tool): Tool {
  const locked = (GUARDED_PARAMS[definition.name] ?? []).filter(guard => !unlocked(guard.env));
  if (locked.length === 0) return definition;
  const properties = { ...(definition.inputSchema.properties ?? {}) };
  for (const guard of locked) delete properties[guard.param];
  return { ...definition, inputSchema: { ...definition.inputSchema, properties } };
}

/** Why a call is refused, or null when it may run. */
export function refusalReason(name: string, args: unknown): string | null {
  if (writesEnabled()) return null;
  if (!READ_TOOLS.has(name)) {
    return `${name} is disabled: this build is read-only (MSTEAMS_ALLOW_WRITES=1 enables write tools).`;
  }
  const values = (args ?? {}) as Record<string, unknown>;
  const used = (GUARDED_PARAMS[name] ?? []).find(guard => values[guard.param] && !unlocked(guard.env));
  return used ? `${name} ${used.param} is disabled in this build (${used.env}=1 enables it).` : null;
}

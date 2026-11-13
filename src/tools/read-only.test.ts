import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { getToolDefinitions, getTool, hasTool, invokeTool, toolCategories } from './registry.js';
import { READ_TOOLS, WRITE_TOOLS, refusalReason } from './read-only.js';

function registeredNames(): string[] {
  vi.stubEnv('MSTEAMS_ALLOW_WRITES', '1');
  const names = getToolDefinitions().map(tool => tool.name);
  vi.stubEnv('MSTEAMS_ALLOW_WRITES', '');
  return names;
}

beforeEach(() => { vi.stubEnv('MSTEAMS_ALLOW_WRITES', ''); vi.stubEnv('MSTEAMS_ALLOW_FORCE_LOGIN', ''); });
afterEach(() => vi.unstubAllEnvs());

describe('read-only classification', () => {
  it('classifies every registered tool exactly once', () => {
    const unclassified = registeredNames().filter(name => READ_TOOLS.has(name) === WRITE_TOOLS.has(name));
    expect(unclassified).toEqual([]);
  });
});

describe('read-only schemas', () => {
  // A parameter upstream adds to a read tool can write (markRead did). Review
  // any change here before updating the snapshot.
  it('lists the reviewed parameters for every read tool', () => {
    const params = Object.fromEntries(getToolDefinitions()
      .map(tool => [tool.name, Object.keys(tool.inputSchema.properties ?? {}).sort()])
      .sort(([a], [b]) => String(a).localeCompare(String(b))));
    expect(params).toMatchSnapshot();
  });
});

describe('read-only mode (default)', () => {
  it('lists only read tools', () => {
    const listed = getToolDefinitions().map(tool => tool.name);
    expect(listed.filter(name => !READ_TOOLS.has(name))).toEqual([]);
    expect(listed.sort()).toEqual(registeredNames().filter(name => READ_TOOLS.has(name)).sort());
  });

  it('hides write tools from lookup', () => {
    expect(getTool('teams_send_message')).toBeUndefined();
    expect(hasTool('teams_send_message')).toBe(false);
    expect(hasTool('teams_search')).toBe(true);
  });

  it('refuses a write tool call before validation', async () => {
    const result = await invokeTool('teams_send_message', {});
    expect(result.success).toBe(false);
    if (!result.success) expect(result.error.code).toBe('ACCESS_DENIED');
  });

  it('strips and refuses markRead on teams_get_thread', async () => {
    const thread = getToolDefinitions().find(tool => tool.name === 'teams_get_thread');
    expect(thread?.inputSchema.properties).not.toHaveProperty('markRead');
    const result = await invokeTool('teams_get_thread', { conversationId: '19:x@thread.v2', markRead: true });
    expect(result.success).toBe(false);
    if (!result.success) expect(result.error.code).toBe('ACCESS_DENIED');
  });
});

describe('MSTEAMS_ALLOW_WRITES=1', () => {
  it('restores every tool and the markRead parameter', () => {
    vi.stubEnv('MSTEAMS_ALLOW_WRITES', '1');
    const listed = getToolDefinitions();
    expect(listed.map(tool => tool.name).sort()).toEqual(toolCategories.flatMap(group => group.tools).sort());
    expect(listed.find(tool => tool.name === 'teams_get_thread')?.inputSchema.properties).toHaveProperty('markRead');
    expect(refusalReason('teams_send_message', {})).toBeNull();
  });
});

describe('write parameters', () => {
  it('allows markRead: false and refuses only a truthy value', () => {
    expect(refusalReason('teams_get_thread', { conversationId: 'x', markRead: false })).toBeNull();
    expect(refusalReason('teams_get_thread', { conversationId: 'x', markRead: true })).not.toBeNull();
  });
});

describe('teams_login forceNew', () => {
  const loginSchema = () => getToolDefinitions().find(tool => tool.name === 'teams_login')?.inputSchema.properties;

  it('is stripped and refused by default', async () => {
    expect(loginSchema()).not.toHaveProperty('forceNew');
    const result = await invokeTool('teams_login', { forceNew: true });
    expect(result.success).toBe(false);
    if (!result.success) expect(result.error.code).toBe('ACCESS_DENIED');
    expect(refusalReason('teams_login', { forceNew: false })).toBeNull();
  });

  it('is unlocked by MSTEAMS_ALLOW_FORCE_LOGIN=1 without enabling writes', () => {
    vi.stubEnv('MSTEAMS_ALLOW_FORCE_LOGIN', '1');
    expect(loginSchema()).toHaveProperty('forceNew');
    expect(refusalReason('teams_login', { forceNew: true })).toBeNull();
    expect(refusalReason('teams_send_message', {})).not.toBeNull();
    expect(refusalReason('teams_get_thread', { conversationId: 'x', markRead: true })).not.toBeNull();
  });

  it('is unlocked by MSTEAMS_ALLOW_WRITES=1', () => {
    vi.stubEnv('MSTEAMS_ALLOW_WRITES', '1');
    expect(refusalReason('teams_login', { forceNew: true })).toBeNull();
  });
});

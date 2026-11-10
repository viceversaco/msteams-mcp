import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { getToolDefinitions, getTool, hasTool, invokeTool } from './registry.js';
import { READ_TOOLS, WRITE_TOOLS } from './read-only.js';

function registeredNames(): string[] {
  vi.stubEnv('MSTEAMS_ALLOW_WRITES', '1');
  const names = getToolDefinitions().map(tool => tool.name);
  vi.stubEnv('MSTEAMS_ALLOW_WRITES', '');
  return names;
}

beforeEach(() => vi.stubEnv('MSTEAMS_ALLOW_WRITES', ''));
afterEach(() => vi.unstubAllEnvs());

describe('read-only classification', () => {
  it('classifies every registered tool exactly once', () => {
    const unclassified = registeredNames().filter(name => READ_TOOLS.has(name) === WRITE_TOOLS.has(name));
    expect(unclassified).toEqual([]);
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
    expect(listed.map(tool => tool.name)).toContain('teams_send_message');
    expect(listed.find(tool => tool.name === 'teams_get_thread')?.inputSchema.properties).toHaveProperty('markRead');
  });
});

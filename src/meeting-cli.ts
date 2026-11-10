#!/usr/bin/env node
/**
 * Read-only meeting CLI: transcripts and recordings as JSON (see cli/meeting.ts).
 *
 * It imports only read and download functions, never the tool registry, so it
 * cannot send, post, react or mark read.
 *
 * It sets process.exitCode and never calls process.exit: on Windows node an
 * exit after Substrate calls trips a libuv assertion (src\win\async.c) and
 * turns a good result into exit 9.
 */

import { runMeetingCommand } from './cli/meeting.js';

const { code, out } = await runMeetingCommand(process.argv.slice(2));
process.stdout.write(JSON.stringify(out, null, 1) + '\n');
process.exitCode = code;

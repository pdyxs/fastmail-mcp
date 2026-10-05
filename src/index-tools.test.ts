import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'url';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';

// Spawns the real server over stdio and lists its tools. ListTools needs no
// credentials, so no network calls are made.
describe('MCP tool surface', () => {
  it('registers get_email_content with a required emailId', async () => {
    const transport = new StdioClientTransport({
      command: process.execPath,
      args: ['--import', 'tsx', fileURLToPath(new URL('./index.ts', import.meta.url))],
      env: { PATH: process.env.PATH ?? '' },
      stderr: 'ignore',
    });
    const client = new Client({ name: 'test', version: '0.0.0' });
    await client.connect(transport);
    try {
      const { tools } = await client.listTools();
      const tool = tools.find(t => t.name === 'get_email_content');
      assert.ok(tool, 'get_email_content not registered');
      assert.deepEqual((tool!.inputSchema as any).required, ['emailId']);
    } finally {
      await client.close();
    }
  });
});

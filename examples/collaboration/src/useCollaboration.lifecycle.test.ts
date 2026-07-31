import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';

const hookSource = readFileSync(new URL('./useCollaboration.ts', import.meta.url), 'utf8');
const appSource = readFileSync(new URL('./App.tsx', import.meta.url), 'utf8');

describe('collaboration lifecycle contract', () => {
  test('creates the Yjs document and provider in an effect, not during render', () => {
    expect(hookSource).toMatch(/useEffect\(\(\) => \{[\s\S]*const ydoc = new Y\.Doc\(\);/);
    expect(hookSource).toMatch(/useEffect\(\(\) => \{[\s\S]*new WebrtcProvider/);
    expect(hookSource).not.toMatch(/useMemo\(\(\) => \{[\s\S]*new WebrtcProvider/);
  });

  test('guards DocxEditor mount until collaboration plugins are ready', () => {
    expect(hookSource).toMatch(/ready:\s*boolean/);
    expect(hookSource).toMatch(/return \{[\s\S]*ready:/);
    expect(appSource).toMatch(/if \(!ready\)/);
    expect(appSource).toContain('<DocxEditor');
  });

  test('allows deployments to provide reliable signaling servers', () => {
    expect(hookSource).toContain('VITE_Y_WEBRTC_SIGNALING');
    expect(hookSource).toContain('wss://y-webrtc-eu.fly.dev');
    expect(hookSource).not.toContain('y-webrtc-signaling-eu.herokuapp.com');
  });

  test('hydrates the visible connection status after subscribing', () => {
    expect(hookSource).toContain('provider.connected');
  });

  test('passes sharded comment id allocation into the editor', () => {
    expect(hookSource).toContain('commentIdAllocatorOptions');
    expect(hookSource).toContain('shardStride');
    expect(appSource).toContain('commentIdAllocatorOptions={commentIdAllocatorOptions}');
  });

  test('keeps the local user out of the title bar avatar stack to avoid duplicate awareness labels', () => {
    expect(appSource).toContain('remoteUsers');
    expect(appSource).toContain('users.filter((collaborator) => !collaborator.isLocal)');
    expect(appSource).toContain('<AvatarStack users={remoteUsers} />');
    expect(appSource).not.toContain('<AvatarStack users={users} />');
  });
});

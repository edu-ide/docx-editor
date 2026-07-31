# docx-editor — Realtime Collaboration

A minimal demo of multi-user collaborative editing on top of `@eigenpal/docx-editor-react` using [Yjs](https://yjs.dev), [`y-prosemirror`](https://github.com/yjs/y-prosemirror), and [`y-webrtc`](https://github.com/yjs/y-webrtc).

**No backend required.** Peers find each other through Yjs's public WebRTC signaling servers and sync directly browser-to-browser.

## Try it

```bash
bun install
bun run dev
```

Open <http://localhost:5273>, then click **Share link** and paste the URL into a second browser window. Type in either window — edits, selections, and avatars sync live.

## How it works

Four pieces:

1. **`externalContent` prop** tells the editor to treat its `document` prop as a schema seed only and skip the mount-time content load. `ySyncPlugin` populates ProseMirror from the shared `Y.Doc` instead.
2. **`externalPlugins`** receives `ySyncPlugin`, `yCursorPlugin(awareness)`, and `yUndoPlugin()` so Yjs owns the document state, remote cursors, and history. **Tracked changes sync automatically** through this — `insertion`/`deletion` mark attrs (author, date, revision id) ride along with the synced PM tree. **Remote cursors and selection-range highlights** also surface automatically via the editor's PM-decoration forwarding layer.
3. **Awareness** (Yjs's ephemeral state channel) carries each user's name, color, and selection. The `AvatarStack` in the title bar reads `provider.awareness.getStates()` and renders connected users.
4. **Controlled `comments` prop** + a `Y.Map<Comment>` on the same `Y.Doc`. PM only carries the comment range markers; the thread metadata (text, author, replies, resolved status) lives in the Y.Map, mirrored into React state and merged back through `onCommentsChange`.

```tsx
<DocxEditor
  document={createEmptyDocument()} // schema seed only
  externalContent // skip the load — Yjs owns content
  externalPlugins={[ySync, yCursor, yUndo]}
  comments={comments} // mirrored from Y.Map<Comment>
  onCommentsChange={setComments} // merges back into Y.Map by comment id
  author={user.name} // attribution for comments / track changes
  renderTitleBarRight={() => <AvatarStack users={users} />}
/>
```

## Files

| File                      | What it does                                                                                      |
| ------------------------- | ------------------------------------------------------------------------------------------------- |
| `src/App.tsx`             | Wires identity, room, collaboration hook, and renders `DocxEditor`                                |
| `src/useCollaboration.ts` | Sets up `Y.Doc`, `WebrtcProvider`, awareness, the y-prosemirror plugins, and the comments Y.Map |
| `src/AvatarStack.tsx`     | Overlapping circular avatars in the title bar                                                     |
| `src/identity.ts`         | Per-tab user identity (sessionStorage) and room id from URL hash                                  |

## Caveats

This is a **demo**, not a production-ready collab template. Known gaps:

- **Comment thread conflicts are last-writer-wins per comment id.** Distinct comments are keyed in a `Y.Map`, so two peers adding different comments no longer replace the whole collection. If two peers edit the same comment body at the same instant, the later map value wins; use per-field or rich-text CRDTs inside each thread if you need character-level merging inside comment text.
- **Comment IDs are sharded per peer.** The demo passes `commentIdAllocatorOptions` into `DocxEditor`, so peers allocate from different numeric ID lanes instead of all starting at `1`. Existing document IDs still win because the allocator seeds above loaded comments and tracked-change revision marks.
- **Tracked-change accept/reject races.** Two peers accepting/rejecting the same change at the same instant produce two PM transactions over overlapping ranges. Yjs picks an ordering and the loser's intent is silently dropped — no conflict UI.
- **The public WebRTC signaling servers are best-effort.** For a stable connection, deploy [`y-websocket`](https://github.com/yjs/y-websocket), [PartyKit](https://www.partykit.io/), [Liveblocks](https://liveblocks.io/), or [Hocuspocus](https://tiptap.dev/hocuspocus).
- **Sessions are ephemeral.** Refresh in an empty room → the document disappears. Add [`y-indexeddb`](https://github.com/yjs/y-indexeddb) for local persistence, or a server-side persistence layer for shared persistence.
- **Loading an existing `.docx` into a live room is non-trivial.** The source-of-truth swap from `document`/`documentBuffer` to the `Y.Doc` needs to happen exactly once and only on a designated peer. Out of scope for this demo.

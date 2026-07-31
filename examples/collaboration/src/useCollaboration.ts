import { useCallback, useEffect, useRef, useState } from 'react';
import * as Y from 'yjs';
import { WebrtcProvider } from 'y-webrtc';
import { ySyncPlugin, yCursorPlugin, yUndoPlugin } from 'y-prosemirror';
import type { Plugin } from 'prosemirror-state';
import type { Comment } from '@eigenpal/docx-editor-core/types/content';
import type { CommentIdAllocatorOptions } from '@eigenpal/docx-editor-core/prosemirror/commentIdAllocator';
import { applyCommentSnapshotToMap, commentsFromMap } from './sharedComments';

export interface CollaborativeUser {
  clientId: number;
  name: string;
  color: string;
  isLocal: boolean;
}

export interface CollaborationState {
  ready: boolean;
  plugins: Plugin[];
  users: CollaborativeUser[];
  roomName: string;
  status: 'connecting' | 'connected' | 'disconnected';
  /** Comments mirrored from a Y.Map on the same Y.Doc — pass to DocxEditor's `comments` prop. */
  comments: Comment[];
  /** Pass to DocxEditor's `onCommentsChange`. Merges comment changes by stable id. */
  setComments: (next: Comment[]) => void;
  /** Sharded comment/revision id allocation for this peer. */
  commentIdAllocatorOptions: CommentIdAllocatorOptions;
}

const DEFAULT_SIGNALING_SERVERS = ['wss://y-webrtc-eu.fly.dev'];
const COMMENT_ID_SHARD_STRIDE = 1_000_000;

function commentIdAllocatorOptionsForClient(clientId: number): CommentIdAllocatorOptions {
  return {
    shardOffset: (Math.abs(Math.floor(clientId)) % COMMENT_ID_SHARD_STRIDE) + 1,
    shardStride: COMMENT_ID_SHARD_STRIDE,
  };
}

function getSignalingServers(): string[] {
  const configured = (
    import.meta as unknown as { env?: { VITE_Y_WEBRTC_SIGNALING?: string } }
  ).env?.VITE_Y_WEBRTC_SIGNALING;
  const servers = configured
    ?.split(',')
    .map((server) => server.trim())
    .filter(Boolean);
  return servers?.length ? servers : DEFAULT_SIGNALING_SERVERS;
}

interface CollaborationRuntime {
  roomName: string;
  ydoc: Y.Doc;
  provider: WebrtcProvider;
  plugins: Plugin[];
  yComments: Y.Map<Comment>;
  commentIdAllocatorOptions: CommentIdAllocatorOptions;
}

export function useCollaboration(
  roomName: string,
  localUser: { name: string; color: string }
): CollaborationState {
  const [runtime, setRuntime] = useState<CollaborationRuntime | null>(null);
  const [users, setUsers] = useState<CollaborativeUser[]>([]);
  const [status, setStatus] = useState<'connecting' | 'connected' | 'disconnected'>('connecting');
  const [comments, setCommentsState] = useState<Comment[]>([]);
  const commentsRef = useRef<Comment[]>([]);
  const activeRuntime = runtime?.roomName === roomName ? runtime : null;

  // Y.Doc, provider, prosemirror plugins, and the comments Y.Map are created
  // in an effect: the provider joins the room as a side effect, and aborted
  // renders must not leak a connection.
  useEffect(() => {
    setUsers([]);
    commentsRef.current = [];
    setCommentsState([]);
    setStatus('connecting');

    const ydoc = new Y.Doc();
    const provider = new WebrtcProvider(roomName, ydoc, { signaling: getSignalingServers() });
    const fragment = ydoc.getXmlFragment('prosemirror');
    const plugins = [ySyncPlugin(fragment), yCursorPlugin(provider.awareness), yUndoPlugin()];
    const yComments = ydoc.getMap<Comment>('comments');
    const commentIdAllocatorOptions = commentIdAllocatorOptionsForClient(provider.awareness.clientID);

    setRuntime({ roomName, ydoc, provider, plugins, yComments, commentIdAllocatorOptions });

    return () => {
      provider.destroy();
      ydoc.destroy();
    };
  }, [roomName]);

  // Publish local user identity into awareness so peers can render avatars + cursors.
  useEffect(() => {
    if (!activeRuntime) return;
    activeRuntime.provider.awareness.setLocalStateField('user', localUser);
  }, [activeRuntime, localUser.name, localUser.color]);

  // Subscribe to awareness + connection changes.
  useEffect(() => {
    if (!activeRuntime) return;

    const { provider } = activeRuntime;
    const refreshUsers = () => {
      const localId = provider.awareness.clientID;
      const all: CollaborativeUser[] = [];
      provider.awareness.getStates().forEach((state, clientId) => {
        if (!state.user) return;
        all.push({
          clientId,
          name: state.user.name,
          color: state.user.color,
          isLocal: clientId === localId,
        });
      });
      setUsers(all);
    };
    const handleStatus = (event: { connected: boolean }) => {
      setStatus(event.connected ? 'connected' : 'disconnected');
    };

    refreshUsers();
    provider.awareness.on('change', refreshUsers);
    provider.on('status', handleStatus);
    handleStatus({ connected: provider.connected });

    return () => {
      provider.awareness.off('change', refreshUsers);
      provider.off('status', handleStatus);
    };
  }, [activeRuntime]);

  // Mirror the Y.Map<Comment> into React state. Each comment is keyed by its
  // stable Word id, so concurrent peers can add or edit different threads
  // without replacing the whole collection.
  useEffect(() => {
    if (!activeRuntime) return;
    const { yComments } = activeRuntime;
    const sync = () => {
      const next = commentsFromMap(yComments);
      commentsRef.current = next;
      setCommentsState(next);
    };
    sync();
    yComments.observe(sync);
    return () => yComments.unobserve(sync);
  }, [activeRuntime]);

  // Push the editor's new comments array back into Yjs. We diff against the
  // local snapshot this client has actually seen; comments concurrently added
  // by another peer are not deleted merely because this callback received a
  // stale array.
  const setComments = useCallback(
    (next: Comment[]) => {
      if (!activeRuntime) return;
      const { ydoc, yComments } = activeRuntime;
      applyCommentSnapshotToMap(ydoc, yComments, commentsRef.current, next);
    },
    [activeRuntime]
  );

  return {
    ready: Boolean(activeRuntime),
    plugins: activeRuntime?.plugins ?? [],
    users: activeRuntime ? users : [],
    roomName,
    status: activeRuntime ? status : 'connecting',
    comments: activeRuntime ? comments : [],
    setComments,
    commentIdAllocatorOptions: activeRuntime?.commentIdAllocatorOptions ?? {},
  };
}

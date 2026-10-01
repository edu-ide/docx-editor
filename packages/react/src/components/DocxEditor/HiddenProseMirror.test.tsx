import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, mock, test } from 'bun:test';
import { act, cleanup, render } from '@testing-library/react';
import { createRef, StrictMode } from 'react';
import { Plugin } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { HiddenProseMirror, type HiddenProseMirrorRef } from './HiddenProseMirror';

beforeAll(() => GlobalRegistrator.register());
afterEach(() => cleanup());
afterAll(() => GlobalRegistrator.unregister());

describe('HiddenProseMirror transaction lifecycle', () => {
  test('ignores a queued transaction after unmount without reviving the view ref', async () => {
    const ref = createRef<HiddenProseMirrorRef>();
    const onTransaction = mock(() => {});
    const onSelectionChange = mock(() => {});
    const mounted = render(
      <HiddenProseMirror
        ref={ref}
        document={null}
        onTransaction={onTransaction}
        onSelectionChange={onSelectionChange}
      />
    );
    const handle = ref.current!;
    const view = handle.getView()!;
    const state = view.state;
    const transaction = state.tr.insertText('queued edit');
    const queuedDispatch = Promise.resolve().then(() => view.dispatch(transaction));

    mounted.unmount();
    expect(view.isDestroyed).toBe(true);
    await queuedDispatch;

    expect(handle.getView()).toBeNull();
    expect(view.state).toBe(state);
    expect(onTransaction).not.toHaveBeenCalled();
    expect(onSelectionChange).not.toHaveBeenCalled();
  });

  test('rejects an old view transaction after StrictMode recreates the live view', async () => {
    const ref = createRef<HiddenProseMirrorRef>();
    const views: EditorView[] = [];
    const onTransaction = mock(() => {});
    render(
      <StrictMode>
        <HiddenProseMirror
          ref={ref}
          document={null}
          onEditorViewReady={(view) => views.push(view)}
          onTransaction={onTransaction}
        />
      </StrictMode>
    );
    expect(views).toHaveLength(2);
    const [oldView, liveView] = views;
    const oldState = oldView.state;
    expect(oldView.isDestroyed).toBe(true);
    expect(liveView.isDestroyed).toBe(false);

    await Promise.resolve().then(() => oldView.dispatch(oldState.tr.insertText('stale edit')));

    expect(oldView.state).toBe(oldState);
    expect(ref.current!.getView()).toBe(liveView);
    expect(onTransaction).not.toHaveBeenCalled();
    act(() => liveView.dispatch(liveView.state.tr.insertText('live edit')));
    expect(liveView.state.doc.textContent).toBe('live edit');
    expect(onTransaction).toHaveBeenCalledTimes(1);
  });

  test('still accepts plugin transactions during EditorView construction', () => {
    const ref = createRef<HiddenProseMirrorRef>();
    const onTransaction = mock(() => {});
    const constructionPlugin = new Plugin({
      view(view) {
        view.dispatch(view.state.tr.insertText('initial edit'));
        return {};
      },
    });
    render(
      <HiddenProseMirror
        ref={ref}
        document={null}
        externalPlugins={[constructionPlugin]}
        onTransaction={onTransaction}
      />
    );

    expect(ref.current!.getView()!.isDestroyed).toBe(false);
    expect(ref.current!.getState()!.doc.textContent).toBe('initial edit');
    expect(onTransaction).toHaveBeenCalledTimes(1);
  });
});

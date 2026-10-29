/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import '../../../../react/test/dom-setup.ts';
(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
import { afterEach, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render } from '@testing-library/react';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { DocxEditorRoot } from '../../../../react/src/editor/DocxEditorRoot.tsx';
import { DocxEditorViewport } from '../../../../react/src/editor/DocxEditorViewport.tsx';
import { DocxEditorContent } from '../../../../react/src/editor/DocxEditorContent.tsx';
import { DocxEditorHyperLink } from '../../../../react/src/editor/DocxEditorHyperLink.tsx';
import { ContextMenu } from '../../../../react/src/editor/contextmenu/index.ts';
import { collaborationModule } from '../collaboration-module.ts';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support.ts';
import { packageFingerprint, saveReopenDigest } from './document-support.ts';

const harness = createPeerHarness('native-input-menu');
afterEach(() => {
  cleanup();
  harness.cleanup();
});

test('native popup menus preserve two collaborating editors and peer edits', async () => {
  const pair = await harness.pair(zipDocument('<w:p><w:r><w:t>hello world</w:t></w:r></w:p>'));
  const mount = (peer: Peer) => {
    peer.detach();
    let editor: DocxEditorInstance | null = null;
    const view = render(
      <DocxEditorRoot
        document={peer.room.document}
        modules={[collaborationModule({ session: peer.room.session })]}
        onReady={(ready) => {
          editor = ready as DocxEditorInstance;
        }}
      >
        <DocxEditorViewport>
          <DocxEditorContent />
          <DocxEditorHyperLink />
          <ContextMenu />
        </DocxEditorViewport>
      </DocxEditorRoot>
    );
    return { view, editor: editor! };
  };
  const alice = mount(pair.alice),
    bob = mount(pair.bob);
  const paragraphId = alice.editor.surface!.session.paragraphIds()[0]!;
  const selection = { anchor: { paragraphId, offset: 0 }, head: { paragraphId, offset: 5 } };
  act(() => alice.editor.surface!.setSelection(selection));
  const before = packageFingerprint(alice.editor.surface!.session.currentPackage());
  act(() =>
    fireEvent.keyDown(alice.view.container.querySelector('.docx-pages')!, {
      key: 'k',
      metaKey: true,
    })
  );
  const input = alice.view.getByTestId('hyperlink-popup-url-input') as HTMLInputElement;
  input.focus();
  input.value = 'https://example.com';
  const event = new MouseEvent('contextmenu', {
    bubbles: true,
    cancelable: true,
    button: 2,
    clientX: 120,
    clientY: 140,
  });
  act(() => {
    input.dispatchEvent(event);
  });
  expect(event.defaultPrevented).toBe(false);
  expect(document.activeElement).toBe(input);
  expect(alice.view.container.querySelector('.docx-contextmenu')).toBeNull();
  expect(alice.editor.surface!.state().selection).toEqual(selection);
  expect(packageFingerprint(alice.editor.surface!.session.currentPackage())).toBe(before);
  expect(packageFingerprint(bob.editor.surface!.session.currentPackage())).toBe(before);
  act(() => {
    bob.editor.surface!.setSelection({
      anchor: { paragraphId, offset: 6 },
      head: { paragraphId, offset: 11 },
    });
    bob.editor.surface!.type('peer');
    pair.bob.room.session.flushPendingJournals();
    pair.alice.room.session.flushPendingJournals();
  });
  expect(alice.editor.surface!.session.bodyText()).toBe('hello peer');
  const select = (editor: DocxEditorInstance, start: number, end = start) =>
    editor.surface!.setSelection({
      anchor: { paragraphId, offset: start },
      head: { paragraphId, offset: end },
    });
  const converge = () => {
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
    expect(packageFingerprint(alice.editor.surface!.session.currentPackage())).toBe(
      packageFingerprint(bob.editor.surface!.session.currentPackage())
    );
    expect(saveReopenDigest(alice.editor.surface!.session.currentPackage())).toEqual(
      saveReopenDigest(bob.editor.surface!.session.currentPackage())
    );
  };
  act(() => {
    pair.pause();
    select(alice.editor, 0);
    select(bob.editor, 10);
    alice.editor.surface!.type('A');
    bob.editor.surface!.type('B');
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
    pair.resume();
  });
  converge();
  expect(alice.editor.surface!.session.bodyText()).toBe('Ahello peerB');
  act(() => expect(alice.editor.exec({ type: 'undo' }).ok).toBe(true));
  converge();
  expect(bob.editor.surface!.session.bodyText()).toBe('hello peerB');
  act(() => expect(alice.editor.exec({ type: 'redo' }).ok).toBe(true));
  converge();
  expect(alice.editor.surface!.session.bodyText()).toBe('Ahello peerB');
  act(() => {
    pair.pause();
    select(alice.editor, 7, 11);
    select(bob.editor, 11);
    alice.editor.surface!.deleteBackward();
    bob.editor.surface!.type('C');
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
    pair.resume();
  });
  converge();
  expect(alice.editor.surface!.session.bodyText()).toBe('Ahello CB');
  expect(packageFingerprint(alice.editor.surface!.session.currentPackage())).toBe(
    packageFingerprint(bob.editor.surface!.session.currentPackage())
  );
  expect(saveReopenDigest(alice.editor.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(bob.editor.surface!.session.currentPackage())
  );
});

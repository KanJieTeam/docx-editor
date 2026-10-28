/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import '../../../../react/test/dom-setup.ts';

import { afterEach, expect, test } from 'bun:test';
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { Translations } from '@docx-editor.dev/i18n';
import { DocxEditorRoot } from '../../../../react/src/editor/DocxEditorRoot.tsx';
import { DocxEditorViewport } from '../../../../react/src/editor/DocxEditorViewport.tsx';
import { DocxEditorContent } from '../../../../react/src/editor/DocxEditorContent.tsx';
import { DocxEditorToolbar } from '../../../../react/src/editor/toolbar/index.ts';
import { LocaleProvider } from '../../../../react/src/i18n/LocaleContext.tsx';
import {
  SOURCE,
  ITALIAN,
  FRENCH,
  STYLES,
} from '../../../../../scripts/test/paragraph-style-label-fixture.ts';
import { collaborationModule } from '../collaboration-module.ts';
import { createPeerHarness, nodeText, type Peer } from './document-peer-support.ts';
import { packageFingerprint, saveReopenDigest } from './document-support.ts';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const harness = createPeerHarness('style-label-locales', { offlineEditing: true });
const mounted: { root: Root; container: HTMLElement }[] = [];
const reopenedEditors: DocxEditorInstance[] = [];

afterEach(async () => {
  await act(async () => {
    for (const view of mounted.splice(0)) {
      view.root.unmount();
      view.container.remove();
    }
  });
  for (const editor of reopenedEditors.splice(0)) editor.destroy();
  harness.cleanup();
});

async function mountPeer(peer: Peer, catalog?: Translations) {
  peer.detach();
  const container = document.createElement('div');
  document.body.append(container);
  const root = createRoot(container);
  mounted.push({ root, container });
  const modules = [collaborationModule({ session: peer.room.session })];
  let editor: DocxEditorInstance | undefined;
  const render = async (i18n?: Translations) => {
    await act(async () => {
      root.render(
        <LocaleProvider i18n={i18n}>
          <DocxEditorRoot
            document={peer.room.document}
            modules={modules}
            onReady={(instance) => {
              editor = instance as DocxEditorInstance;
            }}
          >
            <DocxEditorToolbar preset={false} overflow={false}>
              <DocxEditorToolbar.StylePicker />
            </DocxEditorToolbar>
            <DocxEditorViewport>
              <DocxEditorContent />
            </DocxEditorViewport>
          </DocxEditorRoot>
        </LocaleProvider>
      );
    });
  };
  await render(catalog);
  if (!editor) throw new Error('The collaboration editor is not ready');
  return {
    container,
    get editor(): DocxEditorInstance {
      if (!editor) throw new Error('The collaboration editor is not ready');
      return editor;
    },
    render,
  };
}

function select(editor: DocxEditorInstance, offset: number, end = offset) {
  const paragraphId = editor.surface!.session.paragraphIds()[0]!;
  editor.surface!.setSelection({
    anchor: { paragraphId, offset },
    head: { paragraphId, offset: end },
  });
}

function trigger(container: HTMLElement): HTMLButtonElement {
  const found = container.querySelector<HTMLButtonElement>('.docx-toolbar__style-trigger');
  if (!found) throw new Error('The style picker trigger is missing');
  return found;
}

function packageOf(editor: DocxEditorInstance) {
  return editor.surface!.session.currentPackage();
}

function textOf(editor: DocxEditorInstance) {
  const pkg = packageOf(editor);
  return nodeText(pkg.parts.get(pkg.mainDocumentPart)!.root);
}

function unchangedStyleNames(editor: DocxEditorInstance) {
  const styles = editor.getDocumentStyles();
  for (const [id, name] of STYLES) {
    expect(styles.find((style) => style.styleId === id)?.name).toBe(name);
  }
}

test('style labels stay local while real collaboration sessions share style IDs and edits', async () => {
  const peers = await harness.pair(SOURCE);
  const alice = await mountPeer(peers.alice, ITALIAN);
  const bob = await mountPeer(peers.bob);
  await act(async () => {
    select(alice.editor, 0);
    select(bob.editor, 0);
  });
  const converged = async () => {
    await act(async () => {
      peers.alice.room.session.flushPendingJournals();
      peers.bob.room.session.flushPendingJournals();
      alice.editor.surface!.flushPendingInput();
      bob.editor.surface!.flushPendingInput();
      await new Promise<void>((resolve) =>
        requestAnimationFrame(() => requestAnimationFrame(() => resolve()))
      );
    });
    expect(packageFingerprint(packageOf(alice.editor))).toBe(
      packageFingerprint(packageOf(bob.editor))
    );
    expect(saveReopenDigest(packageOf(alice.editor))).toEqual(
      saveReopenDigest(packageOf(bob.editor))
    );
    unchangedStyleNames(alice.editor);
    unchangedStyleNames(bob.editor);
  };
  await converged();
  expect(trigger(alice.container).textContent).toBe('Titolo 1');
  expect(trigger(bob.container).textContent).toBe('Heading 1');
  const before = packageFingerprint(packageOf(alice.editor));
  const originalEditor = alice.editor;
  await alice.render(FRENCH);
  expect(alice.editor).toBe(originalEditor);
  expect(trigger(alice.container).textContent).toBe('Titre 1');
  expect(trigger(bob.container).textContent).toBe('Heading 1');
  expect(packageFingerprint(packageOf(alice.editor))).toBe(before);
  expect(packageFingerprint(packageOf(bob.editor))).toBe(before);

  await act(async () => {
    trigger(alice.container).click();
  });
  const title = [...alice.container.querySelectorAll<HTMLButtonElement>('[role="option"]')].find(
    (option) => option.textContent?.trim() === 'Titre'
  );
  expect(title).toBeDefined();
  await act(async () => {
    title!.click();
  });
  await converged();
  expect(alice.editor.snapshot().formatting?.styleId).toBe('Title');
  // Check shared metadata independently of remote toolbar snapshot refresh timing.
  const sharedDocument = packageOf(bob.editor);
  expect(JSON.stringify(sharedDocument.parts.get(sharedDocument.mainDocumentPart)!.root)).toContain(
    '"value":"Title"'
  );
  expect(trigger(alice.container).textContent).toBe('Titre');

  peers.pause();
  await act(async () => {
    select(alice.editor, 0);
    expect(alice.editor.exec({ type: 'insertText', text: 'ALICE ' }).ok).toBe(true);
    select(bob.editor, 'Public style text'.length);
    expect(bob.editor.exec({ type: 'insertText', text: ' BOB' }).ok).toBe(true);
    peers.alice.room.session.flushPendingJournals();
    peers.bob.room.session.flushPendingJournals();
    peers.resume();
  });
  await converged();
  expect(textOf(alice.editor)).toContain('ALICE ');
  expect(textOf(alice.editor)).toContain(' BOB');
  await act(async () => {
    expect(alice.editor.exec({ type: 'undo' }).ok).toBe(true);
  });
  await converged();
  expect(textOf(alice.editor)).not.toContain('ALICE ');
  expect(textOf(alice.editor)).toContain(' BOB');
  await act(async () => {
    expect(alice.editor.exec({ type: 'redo' }).ok).toBe(true);
  });
  await converged();

  peers.pause();
  await act(async () => {
    select(alice.editor, 0, 1);
    expect(alice.editor.exec({ type: 'deleteText' }).ok).toBe(true);
    select(bob.editor, textOf(bob.editor).length);
    expect(bob.editor.exec({ type: 'insertText', text: ' REMOTE' }).ok).toBe(true);
    peers.alice.room.session.flushPendingJournals();
    peers.bob.room.session.flushPendingJournals();
    peers.resume();
  });
  await converged();
  expect(textOf(alice.editor)).toContain(' REMOTE');
  const reopened = createDocxEditor({ document: await alice.editor.save() });
  reopenedEditors.push(reopened);
  const mount = document.createElement('div');
  reopened.attach(mount);
  unchangedStyleNames(reopened);
  expect(saveReopenDigest(packageOf(reopened))).toEqual(saveReopenDigest(packageOf(alice.editor)));
  expect(textOf(reopened)).toBe(textOf(alice.editor));
  expect(trigger(alice.container).textContent).toBe('Titre');
  expect(trigger(bob.container).textContent).toBe('Title');
});

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
import { DocxEditorRoot } from '../../../../react/src/editor/DocxEditorRoot';
import { DocxEditorViewport } from '../../../../react/src/editor/DocxEditorViewport';
import { DocxEditorContent } from '../../../../react/src/editor/DocxEditorContent';
import { ToolbarButton } from '../../../../react/src/editor/toolbar/ToolbarButton';
import { LocaleProvider } from '../../../../react/src/i18n';
import { collaborationModule } from '../collaboration-module';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support';
import { packageFingerprint, saveReopenDigest } from './document-support';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
const harness = createPeerHarness('disabled-reason-locales', { offlineEditing: true });
const mounted: { root: Root; container: HTMLElement }[] = [];
const reopened: DocxEditorInstance[] = [];
const italian = {
  _lang: 'it',
  disabledReason: {
    nothingToUndo: 'Nessuna azione da annullare',
    nothingToRedo: 'Nessuna azione da ripetere',
    outermostLevel: 'La selezione è già al livello più esterno',
  },
};
const french = {
  _lang: 'fr',
  disabledReason: {
    nothingToUndo: 'Aucune action à annuler',
    nothingToRedo: 'Aucune action à rétablir',
    outermostLevel: 'La sélection est déjà au niveau le plus externe',
  },
};
const slots = ['history.undo', 'history.redo', 'list.outdent'] as const;

afterEach(async () => {
  await act(async () => {
    for (const view of mounted.splice(0)) {
      view.root.unmount();
      view.container.remove();
    }
  });
  for (const editor of reopened.splice(0)) editor.destroy();
  harness.cleanup();
});
async function mount(peer: Peer, locale: typeof italian) {
  peer.detach();
  const container = document.createElement('div');
  document.body.append(container);
  const root = createRoot(container);
  mounted.push({ root, container });
  const modules = [collaborationModule({ session: peer.room.session })];
  let editor: DocxEditorInstance | undefined;
  const render = async (catalog: typeof italian) => {
    await act(async () =>
      root.render(
        <LocaleProvider i18n={catalog}>
          <DocxEditorRoot
            document={peer.room.document}
            modules={modules}
            onReady={(value) => {
              editor = value as DocxEditorInstance;
            }}
          >
            {slots.map((slot) => (
              <ToolbarButton key={slot} slotId={slot} />
            ))}
            <DocxEditorViewport>
              <DocxEditorContent />
            </DocxEditorViewport>
          </DocxEditorRoot>
        </LocaleProvider>
      )
    );
  };
  await render(locale);
  if (!editor) throw Error('The real collaboration editor is not ready');
  return {
    container,
    get editor(): DocxEditorInstance {
      if (!editor) throw Error('The real collaboration editor is not ready');
      return editor;
    },
    render,
  };
}
function button(container: HTMLElement, slot: string) {
  const found = container.querySelector<HTMLButtonElement>(`[data-slot="${slot}"]`);
  if (!found) throw Error('The toolbar control is missing');
  return found;
}
function select(editor: DocxEditorInstance, paragraph: number, start: number, end = start) {
  const paragraphId = editor.surface!.session.paragraphIds()[paragraph]!;
  editor.surface!.setSelection({
    anchor: { paragraphId, offset: start },
    head: { paragraphId, offset: end },
  });
}
test('peer-local refusal labels preserve shared edits, history, deletion, and save', async () => {
  const peers = await harness.pair(
    zipDocument('<w:p><w:r><w:t>alpha</w:t></w:r></w:p><w:p><w:r><w:t>beta</w:t></w:r></w:p>')
  );
  const alice = await mount(peers.alice, italian),
    bob = await mount(peers.bob, french);
  const empty = (view: typeof alice, locale: typeof italian) => {
    for (const [slot, label] of [
      ['history.undo', locale.disabledReason.nothingToUndo],
      ['history.redo', locale.disabledReason.nothingToRedo],
      ['list.outdent', locale.disabledReason.outermostLevel],
    ]) {
      expect(button(view.container, slot!).disabled).toBe(true);
      expect(button(view.container, slot!).title).toBe(label!);
    }
    const refusal = view.editor.can({ type: 'undo' });
    expect(refusal.ok).toBe(false);
    if (!refusal.ok) expect(refusal.reason).toBe('nothing to undo');
  };
  empty(alice, italian);
  empty(bob, french);
  const original = packageFingerprint(alice.editor.surface!.session.currentPackage());
  await alice.render(french);
  empty(alice, french);
  expect(packageFingerprint(alice.editor.surface!.session.currentPackage())).toBe(original);
  const converged = () => {
    peers.alice.room.session.flushPendingJournals();
    peers.bob.room.session.flushPendingJournals();
    const a = alice.editor.surface!.session.currentPackage(),
      b = bob.editor.surface!.session.currentPackage();
    expect(packageFingerprint(a)).toBe(packageFingerprint(b));
    expect(saveReopenDigest(a)).toEqual(saveReopenDigest(b));
  };
  await act(async () => {
    peers.pause();
    select(alice.editor, 0, 0);
    expect(alice.editor.exec({ type: 'insertText', text: 'A' }).ok).toBe(true);
    select(bob.editor, 1, 0);
    expect(bob.editor.exec({ type: 'insertText', text: 'B' }).ok).toBe(true);
    peers.alice.room.session.flushPendingJournals();
    peers.bob.room.session.flushPendingJournals();
    peers.resume();
  });
  converged();
  await act(async () => {
    expect(alice.editor.exec({ type: 'undo' }).ok).toBe(true);
  });
  converged();
  expect(button(alice.container, 'history.undo').title).toBe(french.disabledReason.nothingToUndo);
  await act(async () => {
    expect(alice.editor.exec({ type: 'redo' }).ok).toBe(true);
  });
  converged();
  expect(button(alice.container, 'history.redo').title).toBe(french.disabledReason.nothingToRedo);
  await act(async () => {
    peers.pause();
    select(alice.editor, 0, 0, 1);
    expect(alice.editor.exec({ type: 'deleteText' }).ok).toBe(true);
    select(bob.editor, 1, 1);
    expect(bob.editor.exec({ type: 'insertText', text: 'X' }).ok).toBe(true);
    peers.alice.room.session.flushPendingJournals();
    peers.bob.room.session.flushPendingJournals();
    peers.resume();
  });
  converged();
  const restored = createDocxEditor({
    container: document.createElement('div'),
    document: await alice.editor.save(),
  });
  reopened.push(restored);
  expect(saveReopenDigest(restored.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(alice.editor.surface!.session.currentPackage())
  );
});

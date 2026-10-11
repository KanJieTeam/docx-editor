/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import { createHash } from 'node:crypto';
import { strFromU8, unzipSync } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { linesOf } from '@docx-editor.dev/core/layout';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import { renderOnlyCollectionFontFixture } from '../../../../core/src/editor/__tests__/render-only-collection-font-fixture.ts';
import { collaborationModule } from '../collaboration-module.ts';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support.ts';

// Happy-dom has no native font loader. This tests the real editor's policy forwarding,
// real HarfBuzz layout and real collaboration sessions, not browser rasterization.
class BrowserFace {
  constructor(
    readonly family: string,
    readonly bytes: ArrayBuffer
  ) {}
  async load(): Promise<unknown> {
    return this;
  }
}

const hash = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');
const enabled = { allowUnsignedDerivedFaces: true, requireSignedFonts: false } as const;
const signedRequired = { allowUnsignedDerivedFaces: true, requireSignedFonts: true } as const;

async function settled(editor: DocxEditorInstance): Promise<void> {
  for (let attempt = 0; attempt < 200; attempt++) {
    if (
      editor.surface &&
      !editor.fontMeasurement().resolving &&
      editor.fontMeasurement().measurer === 'shaped'
    )
      return;
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
  throw new Error('The editor did not complete admitted font shaping');
}

for (const rightPolicy of [enabled, signedRequired, undefined]) {
  test(`render-only collection policy remains local with ${rightPolicy === enabled ? 'enabled' : rightPolicy ? 'signed-required' : 'default'} peer`, async () => {
    const sourceBytes = renderOnlyCollectionFontFixture(),
      sourceHash = hash(sourceBytes);
    const source = {
      id: 'public-dejavu-collection-face1',
      bytes: sourceBytes,
      hash: 'sha256:' + sourceHash,
      faceIndex: 1,
      request: { family: 'DejaVu Sans', weight: 400, style: 'normal' as const },
    };
    const fonts = { sources: [source] };
    const bytes = zipDocument(
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:sz w:val="28"/></w:rPr><w:t>AB</w:t></w:r></w:p>'
    );
    const harness = createPeerHarness('render-only-collection-faces', { offlineEditing: true });
    const views: { editor: DocxEditorInstance; container: HTMLElement }[] = [];
    const faces = new Set<BrowserFace>();
    const oldFace = Object.getOwnPropertyDescriptor(globalThis, 'FontFace');
    const oldFonts = Object.getOwnPropertyDescriptor(document, 'fonts');
    Object.defineProperty(globalThis, 'FontFace', { configurable: true, value: BrowserFace });
    Object.defineProperty(document, 'fonts', {
      configurable: true,
      value: {
        add: (face: BrowserFace) => faces.add(face),
        delete: (face: BrowserFace) => faces.delete(face),
      },
    });
    const mount = (peer: Peer, policy: typeof enabled | typeof signedRequired | undefined) => {
      peer.detach();
      const container = document.createElement('div');
      document.body.append(container);
      const editor = createDocxEditor({
        container,
        document: peer.room.document,
        fonts,
        renderOnlyCollectionPolicy: policy,
        modules: [collaborationModule({ session: peer.room.session })],
      });
      const view = { editor, container };
      views.push(view);
      return view;
    };
    const xml = async (editor: DocxEditorInstance) =>
      strFromU8(unzipSync(new Uint8Array(await editor.save()))['word/document.xml']!);
    const geometry = (editor: DocxEditorInstance) =>
      linesOf(editor.surface!.layout()).map((line) => ({
        box: line.box,
        baseline: line.baseline,
        spans: line.spans.map((span) => ({
          text: span.text,
          range: span.range,
          box: span.box,
          family: span.style.fontFamily,
        })),
      }));
    try {
      const { alice, bob, pause, resume } = await harness.pair(bytes);
      const left = mount(alice, enabled),
        right = mount(bob, rightPolicy);
      await settled(left.editor);
      await settled(right.editor);
      const paragraphId = linesOf(left.editor.surface!.layout())[0]!.spans[0]!.range.paragraphId;
      const select = (editor: DocxEditorInstance, start: number, end = start) =>
        editor.exec({
          type: 'setSelection',
          range: { anchor: { paragraphId, offset: start }, head: { paragraphId, offset: end } },
        });
      const text = () => paragraphTextOf(left.editor.surface!.session.part(), paragraphId);
      const painted = (view: typeof left) =>
        view.container.querySelector<HTMLElement>('.layout-run-text')!.style.fontFamily;
      expect(faces.size).toBe(rightPolicy === enabled ? 2 : 1);
      expect(painted(left)).toContain('docx-embedded-');
      expect(painted(right).includes('docx-embedded-')).toBe(rightPolicy === enabled);
      expect(geometry(left.editor)).toEqual(geometry(right.editor));
      expect(await xml(left.editor)).toBe(await xml(right.editor));
      expect((await xml(left.editor)).includes('docx-embedded-')).toBe(false);
      expect(
        JSON.stringify(alice.awareness.getLocalState()).includes('allowUnsignedDerivedFaces')
      ).toBe(false);
      pause();
      select(left.editor, 1);
      select(right.editor, 2);
      expect(left.editor.exec({ type: 'insertText', text: 'X' }).ok).toBe(true);
      expect(right.editor.exec({ type: 'insertText', text: 'Y' }).ok).toBe(true);
      resume();
      expect(text()).toBe('AXBY');
      expect(await xml(right.editor)).toBe(await xml(left.editor));
      expect(geometry(right.editor)).toEqual(geometry(left.editor));
      expect(left.editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(text()).toBe('ABY');
      expect(left.editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(text()).toBe('AXBY');
      pause();
      select(left.editor, 1, 2);
      select(right.editor, 4);
      expect(left.editor.exec({ type: 'deleteText' }).ok).toBe(true);
      expect(right.editor.exec({ type: 'insertText', text: 'Z' }).ok).toBe(true);
      resume();
      expect(text()).toBe('ABYZ');
      expect(await xml(right.editor)).toBe(await xml(left.editor));
      right.editor.destroy();
      right.container.remove();
      views.splice(views.indexOf(right), 1);
      harness.leave(bob);
      expect(faces.size).toBe(1);
      select(left.editor, 4);
      expect(left.editor.exec({ type: 'insertText', text: 'Q' }).ok).toBe(true);
      const reconnected = mount(await harness.join(alice, 'reconnected'), rightPolicy);
      await settled(reconnected.editor);
      expect(await xml(reconnected.editor)).toBe(await xml(left.editor));
      expect(geometry(reconnected.editor)).toEqual(geometry(left.editor));
      const container = document.createElement('div');
      document.body.append(container);
      const reopened = createDocxEditor({
        container,
        document: new Uint8Array(await left.editor.save()),
        fonts,
        renderOnlyCollectionPolicy: enabled,
      });
      views.push({ editor: reopened, container });
      await settled(reopened);
      expect(await xml(reopened)).toBe(await xml(left.editor));
      expect(geometry(reopened)).toEqual(geometry(left.editor));
      expect(painted({ editor: reopened, container })).toContain('docx-embedded-');
      expect(hash(sourceBytes)).toBe(sourceHash);
      expect(source.faceIndex).toBe(1);
      expect(source.hash).toBe('sha256:' + sourceHash);
      for (const face of faces) {
        expect(face.family).toMatch(/^docx-embedded-/);
        expect(hash(new Uint8Array(face.bytes))).not.toBe(sourceHash);
      }
    } finally {
      for (const view of views) {
        view.editor.destroy();
        view.container.remove();
      }
      harness.cleanup();
      expect(faces.size).toBe(0);
      if (oldFace) Object.defineProperty(globalThis, 'FontFace', oldFace);
      else Reflect.deleteProperty(globalThis, 'FontFace');
      if (oldFonts) Object.defineProperty(document, 'fonts', oldFonts);
      else Reflect.deleteProperty(document, 'fonts');
    }
  });
}

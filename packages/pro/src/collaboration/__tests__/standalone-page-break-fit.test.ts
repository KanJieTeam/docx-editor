/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { linesOf, caretAt } from '@docx-editor.dev/core/layout';
import { collaborationModule } from '../collaboration-module.ts';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support.ts';

const xml = async (editor: DocxEditorInstance) =>
  strFromU8(unzipSync(new Uint8Array(await editor.save()))['word/document.xml']!);
const paragraph = (content: string) =>
  '<w:p><w:pPr>' +
  '<w:spacing w:before="0" w:after="0" w:line="280" w:lineRule="exact"/>' +
  '<w:widowControl w:val="0"/></w:pPr><w:r>' +
  content +
  '</w:r></w:p>';
const fixture = zipDocument(
  '<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/>' +
    '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar>' +
    '</w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
    '<w:tr><w:trPr><w:trHeight w:val="3000" w:hRule="exact"/></w:trPr><w:tc>' +
    paragraph('<w:t>TABLE</w:t>') +
    '</w:tc></w:tr></w:tbl>' +
    paragraph('<w:br w:type="page"/>') +
    paragraph('<w:t>TITLE</w:t>') +
    '<w:sectPr><w:pgSz w:w="4000" w:h="4000"/>' +
    '<w:pgMar w:top="400" w:bottom="400" w:left="400" w:right="400"/></w:sectPr>'
);

test('standalone breaks converge through two-editor edits, undo, deletion, and reconnect', async () => {
  const harness = createPeerHarness('standalone-page-break-fit', { offlineEditing: true });
  const views: { editor: DocxEditorInstance; container: HTMLElement }[] = [];
  const mount = (peer: Peer) => {
    peer.detach();
    const container = document.createElement('div');
    document.body.append(container);
    const editor = createDocxEditor({
      container,
      document: peer.room.document,
      modules: [collaborationModule({ session: peer.room.session })],
    });
    const view = { editor, container };
    views.push(view);
    return view;
  };
  try {
    const { alice, bob, pause, resume } = await harness.pair(fixture);
    const left = mount(alice),
      right = mount(bob);
    const breakId = linesOf(left.editor.surface!.layout()).find((line) =>
      line.spans.some((span) => span.text === '\f')
    )!.range.paragraphId;
    const titleId = linesOf(left.editor.surface!.layout()).find((line) =>
      line.spans.some((span) => span.text === 'TITLE')
    )!.range.paragraphId;
    const select = (editor: DocxEditorInstance, paragraphId: string, start: number, end = start) =>
      editor.exec({
        type: 'setSelection',
        range: { anchor: { paragraphId, offset: start }, head: { paragraphId, offset: end } },
      });
    const converged = async (pages: number) => {
      expect(await xml(left.editor)).toBe(await xml(right.editor));
      expect(left.editor.surface!.layout().pages).toHaveLength(pages);
      expect(right.editor.surface!.layout().pages).toHaveLength(pages);
    };
    await converged(2);
    select(right.editor, breakId, 0);
    expect(
      caretAt(right.editor.surface!.layout(), { paragraphId: breakId, offset: 0 })!.pageIndex
    ).toBe(0);
    select(left.editor, titleId, 2);
    expect(right.editor.surface!.state().selection!.anchor).toEqual({
      paragraphId: breakId,
      offset: 0,
    });
    pause();
    select(left.editor, breakId, 0);
    select(right.editor, titleId, 5);
    expect(left.editor.exec({ type: 'insertText', text: 'X' }).ok).toBe(true);
    expect(right.editor.exec({ type: 'insertText', text: 'B' }).ok).toBe(true);
    resume();
    await converged(3);
    expect(left.editor.exec({ type: 'undo' }).ok).toBe(true);
    await converged(2);
    expect(left.editor.exec({ type: 'redo' }).ok).toBe(true);
    await converged(3);
    pause();
    select(left.editor, breakId, 0, 1);
    select(right.editor, titleId, 6);
    expect(left.editor.exec({ type: 'deleteText' }).ok).toBe(true);
    expect(right.editor.exec({ type: 'insertText', text: 'C' }).ok).toBe(true);
    resume();
    await converged(2);
    right.editor.destroy();
    right.container.remove();
    views.splice(views.indexOf(right), 1);
    harness.leave(bob);
    select(left.editor, titleId, 7);
    expect(left.editor.exec({ type: 'insertText', text: 'D' }).ok).toBe(true);
    const joined = mount(await harness.join(alice, 'reconnected'));
    expect(await xml(joined.editor)).toBe(await xml(left.editor));
    expect(joined.editor.surface!.layout().pages).toHaveLength(2);
    const container = document.createElement('div');
    document.body.append(container);
    const reopened = createDocxEditor({
      container,
      document: new Uint8Array(await left.editor.save()),
    });
    views.push({ editor: reopened, container });
    expect(await xml(reopened)).toBe(await xml(left.editor));
    expect(reopened.surface!.layout().pages).toHaveLength(2);
    expect((await xml(reopened)).includes('w:type="page"')).toBe(true);
    expect((await xml(reopened)).includes('TITLEBCD')).toBe(true);
  } finally {
    for (const view of views) {
      view.editor.destroy();
      view.container.remove();
    }
    harness.cleanup();
  }
});

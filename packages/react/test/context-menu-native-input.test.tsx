import './dom-setup.ts';
(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
import { afterEach, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render } from '@testing-library/react';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { DocxEditorRoot } from '../src/editor/DocxEditorRoot.tsx';
import { DocxEditorViewport } from '../src/editor/DocxEditorViewport.tsx';
import { DocxEditorContent } from '../src/editor/DocxEditorContent.tsx';
import { DocxEditorHyperLink } from '../src/editor/DocxEditorHyperLink.tsx';
import { ContextMenu } from '../src/editor/contextmenu/index.ts';
import { zipSync, strToU8 } from 'fflate';
const SOURCE = zipSync({
  '[Content_Types].xml': strToU8(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  ),
  '_rels/.rels': strToU8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  ),
  'word/document.xml': strToU8(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>hello world</w:t></w:r></w:p></w:body></w:document>'
  ),
});

afterEach(cleanup);
for (const field of ['url', 'textarea'] as const) {
  test(`the ${field} field keeps its native context menu and focus`, () => {
    let editor: DocxEditorInstance | null = null;
    const view = render(
      <DocxEditorRoot
        document={SOURCE}
        onReady={(ready) => {
          editor = ready as DocxEditorInstance;
        }}
      >
        <DocxEditorViewport>
          <DocxEditorContent />
          <DocxEditorHyperLink />
          <ContextMenu />
          <textarea data-testid="host-textarea" />
        </DocxEditorViewport>
      </DocxEditorRoot>
    );
    const instance = editor!;
    const paragraphId = instance.surface!.session.paragraphIds()[0]!;
    act(() =>
      instance.surface!.setSelection({
        anchor: { paragraphId, offset: 0 },
        head: { paragraphId, offset: 5 },
      })
    );
    if (field === 'url')
      act(() =>
        fireEvent.keyDown(view.container.querySelector('.docx-pages')!, { key: 'k', metaKey: true })
      );
    const input = view.getByTestId(
      field === 'url' ? 'hyperlink-popup-url-input' : 'host-textarea'
    ) as HTMLInputElement | HTMLTextAreaElement;
    input.focus();
    input.value = 'https://example.com';
    input.setSelectionRange(0, 5);
    const selection = instance.surface!.state().selection;
    const before = JSON.stringify(instance.surface!.session.part().root);
    const event = new MouseEvent('contextmenu', {
      bubbles: true,
      cancelable: true,
      button: field === 'url' ? 2 : -1,
      clientX: field === 'url' ? 120 : 0,
      clientY: field === 'url' ? 140 : 0,
    });
    act(() => {
      input.dispatchEvent(event);
    });
    expect(event.defaultPrevented).toBe(false);
    expect(view.container.querySelector('.docx-contextmenu')).toBeNull();
    expect(document.activeElement).toBe(input);
    expect(input.value).toBe('https://example.com');
    expect(input.selectionStart).toBe(0);
    expect(input.selectionEnd).toBe(5);
    expect(instance.surface!.state().selection).toEqual(selection);
    expect(JSON.stringify(instance.surface!.session.part().root)).toBe(before);
  });
}

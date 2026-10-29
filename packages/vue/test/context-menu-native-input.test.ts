import './dom-setup.ts';
import { expect, test } from 'bun:test';
import { createApp, h } from 'vue';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { DocxEditorRoot } from '../src/editor/DocxEditorRoot';
import { DocxEditorViewport } from '../src/editor/DocxEditorViewport';
import { DocxEditorContent } from '../src/editor/DocxEditorContent';
import { DocxEditorHyperLink } from '../src/editor/DocxEditorHyperLink';
import { DocxEditorContextMenu } from '../src/editor/contextmenu/index.ts';
import { flush } from './helpers/fixtures.ts';
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

for (const field of ['url', 'textarea'] as const) {
  test(`the ${field} field keeps its native context menu and focus`, async () => {
    const container = document.createElement('div');
    document.body.append(container);
    let editor: DocxEditorInstance | null = null;
    const app = createApp({
      render: () =>
        h(
          DocxEditorRoot,
          {
            document: SOURCE,
            onReady: (ready: unknown) => {
              editor = ready as DocxEditorInstance;
            },
          },
          {
            default: () =>
              h(DocxEditorViewport, null, {
                default: () => [
                  h(DocxEditorContent),
                  h(DocxEditorHyperLink),
                  h(DocxEditorContextMenu),
                  h('textarea', { 'data-testid': 'host-textarea' }),
                ],
              }),
          }
        ),
    });
    app.mount(container);
    try {
      await flush();
      const instance = editor!;
      const paragraphId = instance.surface!.session.paragraphIds()[0]!;
      instance.surface!.setSelection({
        anchor: { paragraphId, offset: 0 },
        head: { paragraphId, offset: 5 },
      });
      if (field === 'url')
        container.querySelector('.docx-pages')!.dispatchEvent(
          new KeyboardEvent('keydown', {
            bubbles: true,
            cancelable: true,
            key: 'k',
            metaKey: true,
          })
        );
      await flush();
      const input = container.querySelector<HTMLInputElement | HTMLTextAreaElement>(
        field === 'url'
          ? '[data-testid="hyperlink-popup-url-input"]'
          : '[data-testid="host-textarea"]'
      )!;
      expect(input).not.toBeNull();
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
      input.dispatchEvent(event);
      await flush();
      expect(event.defaultPrevented).toBe(false);
      expect(container.querySelector('.docx-contextmenu')).toBeNull();
      expect(document.activeElement).toBe(input);
      expect(input.value).toBe('https://example.com');
      expect(input.selectionStart).toBe(0);
      expect(input.selectionEnd).toBe(5);
      expect(instance.surface!.state().selection).toEqual(selection);
      expect(JSON.stringify(instance.surface!.session.part().root)).toBe(before);
    } finally {
      app.unmount();
      container.remove();
    }
  });
}

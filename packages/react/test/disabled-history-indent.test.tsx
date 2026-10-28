import './dom-setup.ts';

(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

import { afterEach, expect, test } from 'bun:test';
import { act, cleanup, render } from '@testing-library/react';
import { zipSync, strToU8 } from 'fflate';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { DocxEditorRoot } from '../src/editor/DocxEditorRoot';
import { DocxEditorViewport } from '../src/editor/DocxEditorViewport';
import { DocxEditorContent } from '../src/editor/DocxEditorContent';
import { ToolbarButton } from '../src/editor/toolbar/ToolbarButton';
import { LocaleProvider } from '../src/i18n';

const SOURCE = zipSync({
  '[Content_Types].xml': strToU8(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  ),
  '_rels/.rels': strToU8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="document" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  ),
  'word/document.xml': strToU8(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:body><w:p><w:r><w:t>Public synthetic text</w:t></w:r></w:p></w:body></w:document>'
  ),
});
const locale = {
  _lang: 'it',
  disabledReason: {
    editorNotReady: 'Editor non pronto',
    nothingToUndo: 'Nessuna azione da annullare',
    nothingToRedo: 'Nessuna azione da ripetere',
    outermostLevel: 'La selezione è già al livello più esterno',
  },
};
const cases = [
  {
    slot: 'history.undo',
    command: { type: 'undo' },
    reason: 'nothing to undo',
    label: locale.disabledReason.nothingToUndo,
  },
  {
    slot: 'history.redo',
    command: { type: 'redo' },
    reason: 'nothing to redo',
    label: locale.disabledReason.nothingToRedo,
  },
  {
    slot: 'list.outdent',
    command: { type: 'adjustIndent', direction: 'decrease' },
    reason: 'the selection is already at the outermost level',
    label: locale.disabledReason.outermostLevel,
  },
] as const;
afterEach(cleanup);

for (const entry of cases)
  test(`${entry.slot} uses the host catalogue without changing the engine refusal`, async () => {
    let editor: DocxEditorInstance | undefined;
    const view = render(
      <LocaleProvider i18n={locale}>
        <DocxEditorRoot
          document={SOURCE}
          onReady={(ready) => {
            editor = ready as DocxEditorInstance;
          }}
        >
          {cases.map((entry) => (
            <ToolbarButton key={entry.slot} slotId={entry.slot} />
          ))}
          <DocxEditorViewport>
            <DocxEditorContent />
          </DocxEditorViewport>
        </DocxEditorRoot>
      </LocaleProvider>
    );
    await act(async () => {});
    expect(editor).toBeDefined();
    const refusal = editor!.can(entry.command);
    expect(refusal.ok).toBe(false);
    if (refusal.ok) throw Error('The synthetic command must be disabled');
    expect(refusal.reason).toBe(entry.reason);
    const button = view.container.querySelector<HTMLButtonElement>(`[data-slot="${entry.slot}"]`)!;
    expect(button.disabled).toBe(true);
    expect(button.title).toBe(entry.label);
  });

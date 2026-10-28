import './dom-setup.ts';
import { afterEach, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render } from '@testing-library/react';
import { hydrateRoot, type Root } from 'react-dom/client';
import { renderToString } from 'react-dom/server';
import { commandForSlotValue, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { readOoxmlPackage, canonicalOoxmlFingerprint } from '@docx-editor.dev/core/store';
import { LocaleProvider } from '../src/i18n/LocaleContext';
import { DocxEditorRoot } from '../src/editor/DocxEditorRoot';
import { DocxEditorViewport } from '../src/editor/DocxEditorViewport';
import { DocxEditorContent } from '../src/editor/DocxEditorContent';
import { DocxEditorToolbar } from '../src/editor/toolbar';
import {
  SOURCE,
  ITALIAN,
  FRENCH,
  STYLES,
} from '../../../scripts/test/paragraph-style-label-fixture';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
const Picker = DocxEditorToolbar.StylePicker;
let editor: DocxEditorInstance;
let hydrated: Root | undefined;
let hydrationContainer: HTMLElement | undefined;
function tree(locale = ITALIAN, t?: (key: string) => string, custom = false) {
  return (
    <LocaleProvider i18n={locale}>
      <DocxEditorRoot
        document={SOURCE}
        onReady={(value) => {
          editor = value as DocxEditorInstance;
        }}
      >
        <DocxEditorToolbar preset={false} overflow={false} t={t}>
          {custom ? (
            <Picker>
              <Picker.Trigger asChild>
                <button data-testid="host-trigger" title="Host title">
                  Host trigger
                </button>
              </Picker.Trigger>
              <Picker.Content>
                <Picker.Item value="Heading2" asChild>
                  <button data-testid="host-item">Host item</button>
                </Picker.Item>
              </Picker.Content>
            </Picker>
          ) : (
            <Picker />
          )}
        </DocxEditorToolbar>
        <DocxEditorViewport>
          <DocxEditorContent />
        </DocxEditorViewport>
      </DocxEditorRoot>
    </LocaleProvider>
  );
}
function trigger(container: HTMLElement): HTMLButtonElement {
  const button = container.querySelector<HTMLButtonElement>('.docx-toolbar__style-trigger');
  if (!button) throw new Error('The style trigger is missing');
  return button;
}
async function select(id: string) {
  const command = commandForSlotValue('styles.style', id);
  if (!command) throw new Error('The style command is missing');
  await act(async () => {
    expect(editor.can(command).ok).toBe(true);
    editor.exec(command);
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
}
async function settle() {
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
}
afterEach(() => {
  if (hydrated) act(() => hydrated!.unmount());
  hydrated = undefined;
  hydrationContainer?.remove();
  hydrationContainer = undefined;
  cleanup();
});

test('built-in style IDs localize both the trigger and their menu entries', async () => {
  const { container } = render(tree());
  await settle();
  expect(trigger(container).textContent).toBe('Titolo 1');
  fireEvent.click(trigger(container));
  const rows = [...container.querySelectorAll<HTMLButtonElement>('[role="option"]')];
  for (const [id, , name] of STYLES) {
    await select(id);
    expect(trigger(container).textContent).toBe(name);
    expect(rows.some((row) => row.textContent?.replace('✓', '').trim() === name)).toBe(true);
  }
});

test('custom name collisions, unmatched ID case, and missing catalog keys keep document labels', async () => {
  const { container } = render(tree());
  await settle();
  for (const [id, name] of [
    ['CustomH1', 'Heading 1'],
    ['heading1', 'Custom case'],
    ['ListParagraph', 'List Paragraph'],
    ['constructor', 'Prototype label'],
  ]) {
    await select(id!);
    expect(trigger(container).textContent).toBe(name!);
  }
});

test('locale and host resolver updates change labels without remounting the editor', async () => {
  const view = render(tree());
  await settle();
  const same = editor;
  fireEvent.click(trigger(view.container));
  view.rerender(tree({ ...ITALIAN, ...FRENCH, styles: { ...ITALIAN.styles, ...FRENCH.styles } }));
  expect(editor).toBe(same);
  expect(trigger(view.container).textContent).toBe('Titre 1');
  expect(
    view.container.querySelector('[role="option"][aria-selected="true"]')!.textContent
  ).toContain('Titre 1');
  view.rerender(tree(ITALIAN, (key) => (key === 'styles.heading1' ? 'Host heading' : key)));
  expect(trigger(view.container).textContent).toBe('Host heading');
  expect(editor.snapshot().formatting?.styleId).toBe('Heading1');
});

test('children and asChild keep host content and commands', async () => {
  const { container } = render(tree(ITALIAN, undefined, true));
  await settle();
  expect(trigger(container).textContent).toBe('Host trigger');
  expect(trigger(container).title).toBe('Host title');
  fireEvent.click(trigger(container));
  const item = container.querySelector<HTMLButtonElement>('[data-testid="host-item"]')!;
  expect(item.textContent).toBe('Host item');
  fireEvent.click(item);
  expect(editor.snapshot().formatting?.styleId).toBe('Heading2');
});

test('localized labels preserve source style definitions and saved paragraph style IDs', async () => {
  const { container } = render(tree());
  await settle();
  await select('Title');
  expect(trigger(container).textContent).toBe('Titolo');
  const source = readOoxmlPackage(SOURCE);
  const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
  expect(source.ok && reopened.ok).toBe(true);
  if (!source.ok || !reopened.ok) throw new Error('The synthetic package must reopen');
  expect(canonicalOoxmlFingerprint(reopened.package.parts.get('/word/styles.xml')!)).toEqual(
    canonicalOoxmlFingerprint(source.package.parts.get('/word/styles.xml')!)
  );
  expect(JSON.stringify(reopened.package.parts.get('/word/document.xml')!.root)).toContain(
    '"value":"Title"'
  );
  expect(JSON.stringify(reopened.package.parts.get('/word/document.xml')!.root)).not.toContain(
    'Titolo'
  );
});

test('server markup hydrates with the same catalog before showing the document style', async () => {
  const markup = renderToString(tree());
  expect(markup).toContain('Testo normale');
  const errors: unknown[] = [];
  hydrationContainer = document.createElement('div');
  hydrationContainer.innerHTML = markup;
  document.body.append(hydrationContainer);
  await act(async () => {
    hydrated = hydrateRoot(hydrationContainer!, tree(), {
      onRecoverableError: (error) => errors.push(error),
    });
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
  expect(errors).toEqual([]);
  expect(trigger(hydrationContainer).textContent).toBe('Titolo 1');
});

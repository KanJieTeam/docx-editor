import './dom-setup.ts';
import { afterEach, expect, test } from 'bun:test';
import { createSSRApp, h, ref, type App } from 'vue';
import { renderToString } from '@vue/server-renderer';
import { commandForSlotValue } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint, readOoxmlPackage } from '@docx-editor.dev/core/store';
import { LocaleProvider } from '../src/i18n/LocaleContext';
import { DocxEditorToolbar } from '../src/editor/toolbar';
import { DocxEditorRoot } from '../src/editor/DocxEditorRoot';
import { DocxEditorViewport } from '../src/editor/DocxEditorViewport';
import { DocxEditorContent } from '../src/editor/DocxEditorContent';
import { mountEditorTree, type MountedEditor } from './helpers/mount';
import { flush } from './helpers/fixtures';
import {
  SOURCE,
  ITALIAN,
  FRENCH,
  STYLES,
} from '../../../scripts/test/paragraph-style-label-fixture';

const Picker = DocxEditorToolbar.StylePicker;
let mounted: MountedEditor | undefined;
let hydrated: App | undefined;
let hydrationContainer: HTMLElement | undefined;
const locale = ref(ITALIAN);
const translate = ref<((key: string) => string) | undefined>();
function toolbar(custom = false) {
  return h(
    DocxEditorToolbar,
    { preset: false, overflow: false, t: translate.value },
    {
      default: () =>
        custom
          ? h(Picker, null, {
              default: () => [
                h(
                  Picker.Trigger,
                  { asChild: true },
                  { default: () => h('button', { title: 'Host title' }, 'Host trigger') }
                ),
                h(Picker.Content, null, {
                  default: () =>
                    h(
                      Picker.Item,
                      { value: 'Heading2', asChild: true },
                      { default: () => h('button', { 'data-testid': 'host-item' }, 'Host item') }
                    ),
                }),
              ],
            })
          : h(Picker),
    }
  );
}
async function mount(custom = false) {
  mounted = mountEditorTree(
    () => toolbar(custom),
    SOURCE,
    undefined,
    undefined,
    undefined,
    (root) => h(LocaleProvider, { i18n: locale.value }, { default: () => root })
  );
  await flush();
  return mounted;
}
function trigger(container: HTMLElement): HTMLButtonElement {
  const button = container.querySelector<HTMLButtonElement>('.docx-toolbar__style-trigger');
  if (!button) throw new Error('The style trigger is missing');
  return button;
}
async function select(id: string) {
  const command = commandForSlotValue('styles.style', id);
  if (!command) throw new Error('The style command is missing');
  expect(mounted!.editor().can(command).ok).toBe(true);
  mounted!.editor().exec(command);
  await flush();
}
afterEach(() => {
  mounted?.unmount();
  mounted = undefined;
  hydrated?.unmount();
  hydrated = undefined;
  hydrationContainer?.remove();
  hydrationContainer = undefined;
  locale.value = ITALIAN;
  translate.value = undefined;
});

test('built-in style IDs localize both the trigger and their menu entries', async () => {
  const view = await mount();
  expect(trigger(view.container).textContent).toBe('Titolo 1');
  trigger(view.container).click();
  await flush();
  for (const [id, , name] of STYLES) {
    await select(id);
    expect(trigger(view.container).textContent).toBe(name);
    expect(
      [...view.container.querySelectorAll('[role="option"]')].some(
        (row) => row.textContent?.replace('✓', '').trim() === name
      )
    ).toBe(true);
  }
});

test('custom name collisions, unmatched ID case, and missing keys retain document labels', async () => {
  const view = await mount();
  for (const [id, name] of [
    ['CustomH1', 'Heading 1'],
    ['heading1', 'Custom case'],
    ['ListParagraph', 'List Paragraph'],
    ['constructor', 'Prototype label'],
  ]) {
    await select(id!);
    expect(trigger(view.container).textContent).toBe(name!);
  }
});

test('locale and host resolver updates keep the same editor and selected style ID', async () => {
  const view = await mount();
  const same = view.editor();
  trigger(view.container).click();
  locale.value = { ...ITALIAN, ...FRENCH, styles: { ...ITALIAN.styles, ...FRENCH.styles } };
  await flush();
  expect(view.editor()).toBe(same);
  expect(trigger(view.container).textContent).toBe('Titre 1');
  expect(
    view.container.querySelector('[role="option"][aria-selected="true"]')!.textContent
  ).toContain('Titre 1');
  translate.value = (key) => (key === 'styles.heading1' ? 'Host heading' : key);
  await flush();
  expect(trigger(view.container).textContent).toBe('Host heading');
  expect(view.editor().snapshot().formatting?.styleId).toBe('Heading1');
});

test('children and asChild retain host content and commands', async () => {
  const view = await mount(true);
  expect(trigger(view.container).textContent).toBe('Host trigger');
  expect(trigger(view.container).title).toBe('Host title');
  trigger(view.container).click();
  await flush();
  const item = view.container.querySelector<HTMLButtonElement>('[data-testid="host-item"]')!;
  expect(item.textContent).toBe('Host item');
  item.click();
  await flush();
  expect(view.editor().snapshot().formatting?.styleId).toBe('Heading2');
});

test('localized labels preserve source style definitions during save and reopen', async () => {
  const view = await mount();
  await select('Title');
  expect(trigger(view.container).textContent).toBe('Titolo');
  const original = readOoxmlPackage(SOURCE);
  const reopened = readOoxmlPackage(new Uint8Array(await view.editor().save()));
  expect(original.ok && reopened.ok).toBe(true);
  if (!original.ok || !reopened.ok) throw new Error('The synthetic package must reopen');
  expect(canonicalOoxmlFingerprint(reopened.package.parts.get('/word/styles.xml')!)).toEqual(
    canonicalOoxmlFingerprint(original.package.parts.get('/word/styles.xml')!)
  );
});

test('server default labels hydrate before the loaded document style appears', async () => {
  const tree = () =>
    h(
      LocaleProvider,
      { i18n: ITALIAN },
      {
        default: () =>
          h(
            DocxEditorRoot,
            { document: SOURCE },
            {
              default: () => [
                toolbar(),
                h(DocxEditorViewport, null, { default: () => h(DocxEditorContent) }),
              ],
            }
          ),
      }
    );
  const markup = await renderToString(createSSRApp({ render: tree }));
  expect(markup).toContain('Testo normale');
  const errors: string[] = [];
  hydrationContainer = document.createElement('div');
  hydrationContainer.innerHTML = markup;
  document.body.append(hydrationContainer);
  hydrated = createSSRApp({ render: tree });
  hydrated.config.warnHandler = (message) => errors.push(message);
  hydrated.mount(hydrationContainer);
  await flush();
  expect(errors).toEqual([]);
  expect(trigger(hydrationContainer).textContent).toBe('Titolo 1');
});

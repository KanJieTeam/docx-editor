import './dom-setup.ts';
import { afterEach, expect, test } from 'bun:test';
import { h } from 'vue';
import { strFromU8, unzipSync } from 'fflate';
import { DocxEditorToolbar } from '../src/editor/toolbar';
import { collaborationModule } from '../../pro/src/collaboration/collaboration-module';
import type { DocumentCollaborationHandle } from '../../pro/src/collaboration/document-session';
import { createPeerHarness } from '../../pro/src/collaboration/__tests__/document-peer-support';
import { flush, mountEditorTree } from './helpers/mount';
import {
  DROPDOWN_SOURCE,
  DROPDOWN_SLOTS,
  POPUP_SELECTOR,
  trackDocumentKeydown,
} from './helpers/dropdown-document';

let unmount: (() => void) | undefined;
afterEach(() => {
  unmount?.();
  unmount = undefined;
});
async function update(action: () => unknown = () => {}) {
  action();
  await flush();
}
async function mount(room?: DocumentCollaborationHandle) {
  const view = mountEditorTree(
    () =>
      h(
        DocxEditorToolbar,
        { preset: false, overflow: false },
        {
          default: () => [
            h(DocxEditorToolbar.Alignment),
            h(DocxEditorToolbar.LineSpacing),
            h(DocxEditorToolbar.TableBorderTarget),
            h(DocxEditorToolbar.TableBorderStyle),
            h(DocxEditorToolbar.TableBorderWidth),
          ],
        }
      ),
    room?.document ?? DROPDOWN_SOURCE,
    () => [],
    room ? [collaborationModule({ session: room.session })] : undefined
  );
  unmount = view.unmount;
  await flush();
  const paragraphId = view.container
    .querySelector('[data-paragraph-id]')!
    .getAttribute('data-paragraph-id')!;
  const caret = { paragraphId, offset: 0 };
  await update(() =>
    view.editor().exec({ type: 'setSelection', range: { anchor: caret, head: caret } })
  );
  return view;
}
const escape = (target: HTMLElement, options: KeyboardEventInit = {}) =>
  target.dispatchEvent(
    new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true, ...options })
  );

test('dropdown dismissal stays local on two collaborating editor instances', async () => {
  const harness = createPeerHarness('vue-dropdown-dismissal', { offlineEditing: true });
  const views: Awaited<ReturnType<typeof mount>>[] = [];
  const xml = async (view: Awaited<ReturnType<typeof mount>>) =>
    strFromU8(unzipSync(new Uint8Array(await view.editor().save()))['word/document.xml']!);
  try {
    const { alice, bob } = await harness.pair(DROPDOWN_SOURCE);
    alice.detach();
    bob.detach();
    views.push(await mount(alice.room));
    views.push(await mount(bob.room));
    const [left, right] = views;
    const trigger = left!.container.querySelector<HTMLButtonElement>(
      '[data-slot="alignment"] [aria-haspopup]'
    )!;
    const other = right!.container.querySelector<HTMLButtonElement>(
      '[data-slot="alignment"] [aria-haspopup]'
    )!;
    const before = await xml(left!);
    await update(() => trigger.click());
    expect(trigger.getAttribute('aria-expanded')).toBe('true');
    expect(other.getAttribute('aria-expanded')).toBe('false');
    await update(() => escape(left!.container.querySelector<HTMLElement>('.docx-pages')!));
    expect(trigger.getAttribute('aria-expanded')).toBe('false');
    expect(await xml(left!)).toBe(before);
    expect(await xml(right!)).toBe(before);
    await update(() => left!.editor().exec({ type: 'insertText', text: 'Z' }));
    expect(await xml(right!)).toBe(await xml(left!));
    expect((await xml(right!)).includes('ZFirst cell')).toBe(true);
    await update(() => left!.editor().exec({ type: 'undo' }));
    expect(await xml(right!)).toBe(before);
    await update(() => left!.editor().exec({ type: 'redo' }));
    expect(await xml(right!)).toBe(await xml(left!));
  } finally {
    for (const view of views) view.unmount();
    await flush();
    unmount = undefined;
    harness.cleanup();
  }
});

for (const slot of DROPDOWN_SLOTS) {
  test(`${slot}: Escape from the document closes the menu without moving focus or changing the document`, async () => {
    const view = await mount();
    const root = view.container.querySelector<HTMLElement>(`[data-slot="${slot}"]`)!;
    expect(root).not.toBeNull();
    const trigger = root.querySelector<HTMLButtonElement>('[aria-haspopup]')!;
    expect(trigger.disabled).toBe(false);
    const pages = view.container.querySelector<HTMLElement>('.docx-pages')!;
    const xml = async () =>
      strFromU8(unzipSync(new Uint8Array(await view.editor().save()))['word/document.xml']!);
    const before = await xml();
    await update(() => {
      pages.focus();
      trigger.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      trigger.click();
    });
    expect(root.querySelector(POPUP_SELECTOR)).not.toBeNull();
    await update(() => {
      pages.focus();
      escape(pages);
    });
    expect(root.querySelector(POPUP_SELECTOR) === null).toBe(true);
    expect(document.activeElement === pages).toBe(true);
    expect(await xml()).toBe(before);
    await update(() => trigger.click());
    expect(root.querySelector(POPUP_SELECTOR)).not.toBeNull();
    await update(() => document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })));
    expect(root.querySelector(POPUP_SELECTOR) === null).toBe(true);
  });

  test(`${slot}: popup Escape returns focus while composition and handled keys stay local`, async () => {
    const view = await mount();
    const root = view.container.querySelector<HTMLElement>(`[data-slot="${slot}"]`)!;
    const trigger = root.querySelector<HTMLButtonElement>('[aria-haspopup]')!;
    await update(() => trigger.click());
    const popup = root.querySelector<HTMLElement>(POPUP_SELECTOR)!;
    const option = popup.querySelector<HTMLButtonElement>('button')!;
    await update(() => {
      option.focus();
      escape(option, { isComposing: true });
    });
    expect(root.querySelector(POPUP_SELECTOR) !== null).toBe(true);
    const handled = (event: Event) => event.preventDefault();
    option.addEventListener('keydown', handled);
    await update(() => escape(option));
    expect(root.querySelector(POPUP_SELECTOR) !== null).toBe(true);
    option.removeEventListener('keydown', handled);
    await update(() => escape(option));
    expect(root.querySelector(POPUP_SELECTOR) === null).toBe(true);
    expect(document.activeElement === trigger).toBe(true);
  });
}

test('composition and a handled Escape do not dismiss a dropdown', async () => {
  const view = await mount();
  const root = view.container.querySelector<HTMLElement>('[data-slot="alignment"]')!;
  const trigger = root.querySelector<HTMLButtonElement>('[aria-haspopup]')!;
  const pages = view.container.querySelector<HTMLElement>('.docx-pages')!;
  await update(() => trigger.click());
  await update(() => escape(pages, { isComposing: true }));
  expect(root.querySelector(POPUP_SELECTOR)).not.toBeNull();
  const event = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true });
  event.preventDefault();
  await update(() => pages.dispatchEvent(event));
  expect(root.querySelector(POPUP_SELECTOR)).not.toBeNull();
  await update(() => escape(pages));
  expect(root.querySelector(POPUP_SELECTOR) === null).toBe(true);
});

test('close, reopen, and unmount remove all document Escape listeners', async () => {
  const view = await mount();
  const root = view.container.querySelector<HTMLElement>('[data-slot="alignment"]')!;
  const trigger = root.querySelector<HTMLButtonElement>('[aria-haspopup]')!;
  const tracked = trackDocumentKeydown(document);
  try {
    await update(() => trigger.click());
    expect(tracked.listeners.size).toBe(1);
    await update(() => escape(view.container.querySelector<HTMLElement>('.docx-pages')!));
    expect(tracked.listeners.size).toBe(0);
    await update(() => trigger.click());
    expect(tracked.listeners.size).toBe(1);
    await update(() => view.unmount());
    unmount = undefined;
    expect(tracked.listeners.size).toBe(0);
  } finally {
    tracked.restore();
  }
});

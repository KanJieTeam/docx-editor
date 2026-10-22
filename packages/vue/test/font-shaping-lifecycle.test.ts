/* eslint-disable react-hooks/rules-of-hooks -- Vue setup uses framework composables. */
import './dom-setup.ts';

import { expect, test } from 'bun:test';
import { createApp, h, nextTick, ref } from 'vue';
import type { Editor } from '@docx-editor.dev/core/contracts/editor';
import type { FontShapingOptions } from '@docx-editor.dev/core';
import {
  scriptGlyphDocument,
  scriptGlyphFonts,
  scriptGlyphGeometry,
  settledScriptGlyphs,
} from '../../core/src/editor/__tests__/script-glyph-fixture.ts';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { DocxEditorRoot } from '../src/editor/DocxEditorRoot';
import { DocxEditorViewport } from '../src/editor/DocxEditorViewport';
import { DocxEditorContent } from '../src/editor/DocxEditorContent';

async function flush(): Promise<void> {
  await nextTick();
  for (let i = 0; i < 10; i++) await new Promise((resolve) => queueMicrotask(resolve));
  await new Promise((resolve) => setTimeout(resolve, 150));
}

test('script mode changes remount with matching geometry; equal modes preserve the editor', async () => {
  const instances: DocxEditorInstance[] = [];
  const mode = ref<FontShapingOptions['scriptGlyphSizing']>(undefined);
  const tick = ref(0);
  const container = document.createElement('div');
  document.body.appendChild(container);
  const app = createApp({
    setup() {
      return () => {
        tick.value;
        return h(
          DocxEditorRoot,
          {
            document: scriptGlyphDocument,
            fonts: scriptGlyphFonts,
            fontShaping: { scriptGlyphSizing: mode.value },
            onReady: (editor: Editor) => instances.push(editor as DocxEditorInstance),
          },
          { default: () => h(DocxEditorViewport, null, { default: () => h(DocxEditorContent) }) }
        );
      };
    },
  });
  try {
    app.mount(container);
    await flush();
    await settledScriptGlyphs(instances[0]!);
    const legacy = scriptGlyphGeometry(instances[0]!);
    expect(legacy.factor).toBe(0.65);
    mode.value = 'legacy';
    await flush();
    expect(instances).toHaveLength(1);
    mode.value = 'font-metrics';
    await flush();
    expect(instances).toHaveLength(2);
    const metric = instances[1]!;
    await settledScriptGlyphs(metric);
    const actual = scriptGlyphGeometry(metric);
    expect(actual.nominal).toBe(14);
    expect(actual.factor).not.toBe(0.65);
    expect(actual.width).not.toBe(legacy.width);
    tick.value++;
    await flush();
    expect(instances).toHaveLength(2);
    expect(scriptGlyphGeometry(metric)).toEqual(actual);
    mode.value = 'legacy';
    await flush();
    expect(instances).toHaveLength(3);
    await settledScriptGlyphs(instances[2]!);
    expect(scriptGlyphGeometry(instances[2]!)).toEqual(legacy);
  } finally {
    app.unmount();
    container.remove();
  }
});

test('equal inline font shaping values preserve edits and undo; changed values remount', async () => {
  const instances: DocxEditorInstance[] = [];
  const tick = ref(0);
  const maxFontBytes = ref<number | undefined>(undefined);
  const container = document.createElement('div');
  document.body.appendChild(container);
  const app = createApp({
    setup() {
      return () => {
        tick.value;
        return h(
          DocxEditorRoot,
          {
            document: 'blank',
            fontShaping: { maxFontBytes: maxFontBytes.value },
            onReady: (editor: Editor) => instances.push(editor as DocxEditorInstance),
          },
          { default: () => h(DocxEditorViewport, null, { default: () => h(DocxEditorContent) }) }
        );
      };
    },
  });
  try {
    app.mount(container);
    await flush();
    expect(instances).toHaveLength(1);
    const first = instances[0]!;
    expect(first.exec({ type: 'insertText', text: 'kept' }).ok).toBe(true);
    const edited = first.surface!.session.bodyText();
    const selection = first.surface!.state().selection;
    tick.value++;
    await flush();
    expect(instances).toHaveLength(1);
    expect(first.surface!.session.bodyText()).toBe(edited);
    expect(first.surface!.state().selection).toEqual(selection);
    expect(first.snapshot().canUndo).toBe(true);
    first.surface!.undo();
    expect(first.surface!.session.bodyText()).not.toBe(edited);
    maxFontBytes.value = 20 * 1024 * 1024;
    await flush();
    expect(instances).toHaveLength(2);
    const second = instances[1]!;
    expect(second).not.toBe(first);
    expect(second.exec({ type: 'insertText', text: 'second' }).ok).toBe(true);
    const secondText = second.surface!.session.bodyText();
    tick.value++;
    await flush();
    expect(instances).toHaveLength(2);
    expect(second.surface!.session.bodyText()).toBe(secondText);
    expect(second.snapshot().canUndo).toBe(true);
  } finally {
    app.unmount();
    container.remove();
  }
});

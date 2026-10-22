import './dom-setup.ts';

(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

import { afterEach, expect, test } from 'bun:test';
import { act, cleanup, render } from '@testing-library/react';
import type { FontShapingOptions } from '@docx-editor.dev/core';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { DocxEditorRoot } from '../src/editor/DocxEditorRoot.tsx';
import { DocxEditorViewport } from '../src/editor/DocxEditorViewport.tsx';
import { DocxEditorContent } from '../src/editor/DocxEditorContent.tsx';
import {
  scriptGlyphDocument,
  scriptGlyphFonts,
  scriptGlyphGeometry,
  settledScriptGlyphs,
} from '../../core/src/editor/__tests__/script-glyph-fixture.ts';

afterEach(cleanup);

test('script mode changes remount with matching geometry; equal modes preserve the editor', async () => {
  const instances: DocxEditorInstance[] = [];
  const tree = (mode: FontShapingOptions['scriptGlyphSizing']) => (
    <DocxEditorRoot
      document={scriptGlyphDocument}
      fonts={scriptGlyphFonts}
      fontShaping={{ scriptGlyphSizing: mode }}
      onReady={(editor) => instances.push(editor as DocxEditorInstance)}
    >
      <DocxEditorViewport>
        <DocxEditorContent />
      </DocxEditorViewport>
    </DocxEditorRoot>
  );
  const view = render(tree(undefined));
  try {
    await act(async () => {
      await settledScriptGlyphs(instances[0]!);
    });
    const first = instances[0]!;
    const legacy = scriptGlyphGeometry(first);
    expect(legacy.factor).toBe(0.65);
    await act(async () => view.rerender(tree('legacy')));
    expect(instances).toHaveLength(1);
    await act(async () => view.rerender(tree('font-metrics')));
    expect(instances).toHaveLength(2);
    const metric = instances[1]!;
    await act(async () => {
      await settledScriptGlyphs(metric);
    });
    const actual = scriptGlyphGeometry(metric);
    expect(actual.nominal).toBe(14);
    expect(actual.factor).not.toBe(0.65);
    expect(actual.width).not.toBe(legacy.width);
    await act(async () => view.rerender(tree('font-metrics')));
    expect(instances).toHaveLength(2);
    expect(scriptGlyphGeometry(metric)).toEqual(actual);
    await act(async () => view.rerender(tree('legacy')));
    expect(instances).toHaveLength(3);
    await act(async () => {
      await settledScriptGlyphs(instances[2]!);
    });
    expect(scriptGlyphGeometry(instances[2]!)).toEqual(legacy);
  } finally {
    view.unmount();
  }
});

test('equal font shaping values preserve edits and undo; changed values remount', async () => {
  const instances: DocxEditorInstance[] = [];
  const tree = (fontShaping?: FontShapingOptions) => (
    <DocxEditorRoot
      document="blank"
      fontShaping={fontShaping}
      onReady={(editor) => instances.push(editor as DocxEditorInstance)}
    >
      <DocxEditorViewport>
        <DocxEditorContent />
      </DocxEditorViewport>
    </DocxEditorRoot>
  );
  const view = render(tree());
  expect(instances).toHaveLength(1);
  const first = instances[0]!;
  await act(async () => {
    expect(first.exec({ type: 'insertText', text: 'kept' }).ok).toBe(true);
  });
  const edited = first.surface!.session.bodyText();
  const selection = first.surface!.state().selection;
  for (const equal of [{}, {}, undefined]) {
    await act(async () => view.rerender(tree(equal)));
    expect(instances).toHaveLength(1);
    expect(first.surface!.session.bodyText()).toBe(edited);
    expect(first.surface!.state().selection).toEqual(selection);
    expect(first.snapshot().canUndo).toBe(true);
  }
  await act(async () => {
    first.surface!.undo();
  });
  expect(first.surface!.session.bodyText()).not.toBe(edited);
  await act(async () => view.rerender(tree({ maxFontBytes: 20 * 1024 * 1024 })));
  expect(instances).toHaveLength(2);
  const second = instances[1]!;
  expect(second).not.toBe(first);
  await act(async () => {
    expect(second.exec({ type: 'insertText', text: 'second' }).ok).toBe(true);
  });
  const secondText = second.surface!.session.bodyText();
  await act(async () => view.rerender(tree({ maxFontBytes: 20 * 1024 * 1024 })));
  expect(instances).toHaveLength(2);
  expect(second.surface!.session.bodyText()).toBe(secondText);
  expect(second.snapshot().canUndo).toBe(true);
});

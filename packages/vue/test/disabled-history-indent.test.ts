import './dom-setup.ts';

import { afterEach, expect, test } from 'bun:test';
import { h } from 'vue';
import { ToolbarButton } from '../src/editor/toolbar/ToolbarButton';
import { LocaleProvider } from '../src/i18n';
import { SOURCE, flush } from './helpers/fixtures';
import { mountEditorTree, type MountedEditor } from './helpers/mount';

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
let mounted: MountedEditor | undefined;
afterEach(() => {
  mounted?.unmount();
  mounted = undefined;
});

for (const entry of cases)
  test(`${entry.slot} uses the host catalogue without changing the engine refusal`, async () => {
    mounted = mountEditorTree(
      () => cases.map((entry) => h(ToolbarButton, { slotId: entry.slot })),
      SOURCE,
      () => [],
      undefined,
      {},
      (root) => h(LocaleProvider, { i18n: locale }, { default: () => [root] })
    );
    await flush();
    const editor = mounted.editor();
    expect(editor).toBeDefined();
    const refusal = editor.can(entry.command);
    expect(refusal.ok).toBe(false);
    if (refusal.ok) throw Error('The synthetic command must be disabled');
    expect(refusal.reason).toBe(entry.reason);
    const button = mounted.container.querySelector<HTMLButtonElement>(
      `[data-slot="${entry.slot}"]`
    )!;
    expect(button.disabled).toBe(true);
    expect(button.title).toBe(entry.label);
  });

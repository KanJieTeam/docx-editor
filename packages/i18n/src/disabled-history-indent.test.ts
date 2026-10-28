import { expect, test } from 'bun:test';
import { createT, deepMerge, en, localizeDisabledReason } from './index';

const italian = {
  ...en,
  disabledReason: {
    ...en.disabledReason,
    nothingToUndo: 'Nessuna azione da annullare',
    nothingToRedo: 'Nessuna azione da ripetere',
    outermostLevel: 'La selezione è già al livello più esterno',
    editorNotReady: 'Editor non pronto',
  },
};
const translate = createT(italian, 'it');
for (const [reason, expected] of [
  ['nothing to undo', italian.disabledReason.nothingToUndo],
  ['nothing to redo', italian.disabledReason.nothingToRedo],
  ['the selection is already at the outermost level', italian.disabledReason.outermostLevel],
]) {
  test(`the host catalogue translates ${reason}`, () => {
    expect(localizeDisabledReason(reason!, translate)).toBe(expected!);
  });
}
test('existing catalogue mappings and exact unknown diagnostics remain available', () => {
  expect(localizeDisabledReason('editor is not ready', translate)).toBe('Editor non pronto');
  expect(localizeDisabledReason('host-specific refusal', translate)).toBe('host-specific refusal');
  expect(localizeDisabledReason(null, translate)).toBeNull();
});
test('English labels remain available when host catalogue entries are null', () => {
  const untranslated = {
    disabledReason: { nothingToUndo: null, nothingToRedo: null, outermostLevel: null },
  };
  const fallback = createT(deepMerge(en, untranslated) as typeof en);
  for (const [reason, expected] of [
    ['nothing to undo', 'There is nothing to undo.'],
    ['nothing to redo', 'There is nothing to redo.'],
    [
      'the selection is already at the outermost level',
      'The selection is already at the outermost level.',
    ],
  ]) {
    expect(localizeDisabledReason(reason!, createT(en))).toBe(expected!);
    expect(localizeDisabledReason(reason!, fallback)).toBe(expected!);
  }
});

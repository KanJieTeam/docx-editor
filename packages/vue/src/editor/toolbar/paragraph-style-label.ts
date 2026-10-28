// Built-in labels belong to the interface; document style names remain unchanged.
const STYLE_LABELS = new Map([
  ['Normal', 'styles.normalText'],
  ['Title', 'styles.title'],
  ['Subtitle', 'styles.subtitle'],
  ['Heading1', 'styles.heading1'],
  ['Heading2', 'styles.heading2'],
  ['Heading3', 'styles.heading3'],
]);

export function paragraphStyleLabel(
  option: { readonly styleId: string; readonly name: string },
  label: (key: string) => string
): string {
  const key = STYLE_LABELS.get(option.styleId);
  return key === undefined ? option.name : label(key);
}

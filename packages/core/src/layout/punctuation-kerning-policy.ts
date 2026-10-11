import type { OoxmlElement } from '@docx-editor.dev/core/store';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// Absent means the default permission. Malformed or misplaced flags turn this
// subset off. They do not reject the document or alter its canonical settings.
export function sourcePunctuationKerningPermission(root: OoxmlElement | null): boolean | undefined {
  if (!root) return undefined;
  if (root.namespaceUri !== W || root.localName !== 'settings') return false;
  const queue: { node: OoxmlElement; depth: number }[] = [{ node: root, depth: 0 }];
  let flag: OoxmlElement | undefined;
  let visited = 0;
  while (queue.length) {
    const item = queue.pop()!;
    if (++visited > 8192 || item.depth > 32) return false;
    for (const child of item.node.children) {
      if (!('localName' in child)) continue;
      if (child.localName === 'noPunctuationKerning') {
        if (flag || item.node !== root || child.namespaceUri !== W) return false;
        flag = child;
      }
      queue.push({ node: child, depth: item.depth + 1 });
    }
  }
  if (!flag) return undefined;
  const values = flag.attributes.filter((a) => a.localName === 'val');
  if (values.length > 1 || values.some((a) => a.namespaceUri !== W)) return false;
  const value = values[0]?.value;
  if (value === undefined || ['1', 'true', 'on'].includes(value)) return false;
  return ['0', 'false', 'off'].includes(value);
}

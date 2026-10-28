// Note reference formatting belongs to the reference, not to the next character.

import { withFreshIds } from '../package/hf-lifecycle-shell.ts';
import { WML_NAMESPACE_URI } from '../package/ooxml-shared.ts';
import { replaceChildren, type EditOptions, type OoxmlEditResult } from '../package/ooxml-edit.ts';
import type { OoxmlNode, OoxmlParagraphNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { NOT_INHERITED } from './mark-character-style-run.ts';
import {
  isRunPropertiesNode,
  paragraphPropertiesNodeOf,
  parentOf,
  runPropertiesNodeOf,
} from './tree-op-nodes.ts';
import { build } from './tree-op-tracked-builders.ts';
import { paragraphOffsetIndex } from './tree-op-segments.ts';

/** The canonical body-side note reference immediately before this caret, if any. */
export function noteReferenceEndingAt(paragraph: OoxmlParagraphNode, offset: number) {
  return paragraphOffsetIndex(paragraph).segments.find(
    (segment) => segment.end === offset && segment.node.kind === 'noteReference'
  );
}

/** Use the paragraph mark's face. Never copy the reference's run properties. */
export function noteTypingProperties(
  paragraph: OoxmlParagraphNode,
  mint: () => string
): OoxmlNode[] {
  const paragraphProperties = paragraphPropertiesNodeOf(paragraph);
  const properties = paragraphProperties ? runPropertiesNodeOf(paragraphProperties) : undefined;
  if (!properties) return [];
  const children = properties.children.filter(
    (node) =>
      node.kind === 'textValue' ||
      node.namespaceUri !== WML_NAMESPACE_URI ||
      !NOT_INHERITED.has(node.localName)
  );
  return children.length ? [withFreshIds({ ...properties, children } as OoxmlNode, mint)] : [];
}

/** Split a shared reference run without changing its reference or remaining text. */
export function insertAfterNoteInRun(
  part: OoxmlPart,
  paragraph: OoxmlParagraphNode,
  run: OoxmlNode,
  index: number,
  nodes: readonly OoxmlNode[],
  mint: () => string,
  options?: EditOptions
): OoxmlEditResult | null {
  const parent = parentOf(part, run.id);
  if (!parent || run.kind !== 'run') return null;
  const head = { ...run, children: run.children.slice(0, index) } as OoxmlNode;
  const remaining = run.children.slice(index);
  const tail = remaining.length
    ? [
        {
          ...run,
          id: mint(),
          children: [
            ...run.children.filter(isRunPropertiesNode).map((node) => withFreshIds(node, mint)),
            ...remaining,
          ],
        } as OoxmlNode,
      ]
    : [];
  const inserted = build(
    mint(),
    'run',
    'r',
    [],
    [...noteTypingProperties(paragraph, mint), ...nodes]
  );
  return replaceChildren(
    part,
    parent.id,
    parent.children.flatMap((child) => (child.id === run.id ? [head, inserted, ...tail] : [child])),
    options
  );
}

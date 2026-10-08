import { useEffect, type RefObject } from 'react';
import { editorScopeFor } from '../editor-scope';

/**
 * The box that holds this editor's chrome and its painted pages.
 *
 * A packaged editor or a wrapped composition answers through `editorScopeFor`. A bare
 * composition has no wrapper, so the nearest ancestor that holds a pages layer stands in.
 */
function editorOwnerFor(root: HTMLElement): Element {
  const scope = editorScopeFor(root);
  if (scope) return scope;
  let box: Element | null = root;
  while (box && !box.querySelector('.docx-pages')) box = box.parentElement;
  return box ?? root;
}

/**
 * Close a toolbar dropdown on an outside press or on Escape.
 *
 * Escape is read on the owner document in the capture phase, so it reaches this listener
 * before the painted surface and before any popup handler. An Escape from inside this
 * editor closes the dropdown and stops there, so the same key does not also leave a
 * header, a note, or the format painter. An Escape from host content or from another
 * editor closes the dropdown and keeps going. A hidden slot closes its dropdown.
 */
export function useDropdownClose(
  open: boolean,
  setOpen: (open: boolean) => void,
  rootRef: RefObject<HTMLElement | null>,
  hidden = false
): void {
  useEffect(() => {
    if (open && hidden) setOpen(false);
  }, [open, hidden, setOpen]);

  useEffect(() => {
    const root = rootRef.current;
    if (!open || hidden || !root) return;
    const doc = root.ownerDocument;
    const inside = (event: Event, box: Element) => event.composedPath().includes(box);
    const onMouseDown = (event: MouseEvent) => {
      if (!inside(event, root)) setOpen(false);
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== 'Escape' || event.isComposing || event.keyCode === 229) return;
      setOpen(false);
      if (!inside(event, root) && !inside(event, editorOwnerFor(root))) return;
      event.preventDefault();
      event.stopPropagation();
      const focused = doc.activeElement;
      if (focused && root.contains(focused)) {
        root.querySelector<HTMLElement>('[aria-haspopup]')?.focus();
      }
    };
    doc.addEventListener('mousedown', onMouseDown);
    doc.addEventListener('keydown', onKeyDown, true);
    return () => {
      doc.removeEventListener('mousedown', onMouseDown);
      doc.removeEventListener('keydown', onKeyDown, true);
    };
  }, [open, hidden, setOpen, rootRef]);
}

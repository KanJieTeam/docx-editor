import { watch, type Ref } from 'vue';

/** Close a mounted dropdown without taking focus from the document. */
export function useDropdownClose(
  open: Ref<boolean>,
  setOpen: (open: boolean) => void,
  rootRef: Ref<HTMLElement | null>
): void {
  watch(
    open,
    (isOpen, _, onCleanup) => {
      const root = rootRef.value;
      if (!isOpen || !root) return;
      const doc = root.ownerDocument;
      const inside = (event: Event) => event.composedPath().includes(root);
      const canDismiss = (event: KeyboardEvent) =>
        event.key === 'Escape' &&
        !event.defaultPrevented &&
        !event.isComposing &&
        event.keyCode !== 229;
      const onMouseDown = (event: MouseEvent) => {
        if (!inside(event)) setOpen(false);
      };
      const onDocumentKeyDown = (event: KeyboardEvent) => {
        // A popup's own keyboard handler runs first in the bubble phase.
        if (!canDismiss(event) || inside(event)) return;
        event.preventDefault();
        setOpen(false);
      };
      const onRootKeyDown = (event: KeyboardEvent) => {
        if (!canDismiss(event)) return;
        event.preventDefault();
        root.querySelector<HTMLElement>('[aria-haspopup]')?.focus();
        setOpen(false);
      };
      doc.addEventListener('mousedown', onMouseDown);
      doc.addEventListener('keydown', onDocumentKeyDown);
      root.addEventListener('keydown', onRootKeyDown);
      onCleanup(() => {
        doc.removeEventListener('mousedown', onMouseDown);
        doc.removeEventListener('keydown', onDocumentKeyDown);
        root.removeEventListener('keydown', onRootKeyDown);
      });
    },
    { flush: 'post' }
  );
}

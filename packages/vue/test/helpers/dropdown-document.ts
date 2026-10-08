import { docx } from './fixtures';

export const DROPDOWN_SOURCE = docx(
  '<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="4800"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
    '<w:tr><w:tc><w:p><w:r><w:t>First cell</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>Second cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
    '<w:p><w:r><w:t>After the table</w:t></w:r></w:p>'
);

export const DROPDOWN_SLOTS = [
  'alignment',
  'list.lineSpacing',
  'table.borderTarget',
  'table.borderStyle',
  'table.borderWidth',
] as const;

export const POPUP_SELECTOR =
  '.docx-toolbar__alignment-popup, .docx-toolbar__line-spacing-menu, .docx-table-chrome__panel';

export function trackDocumentKeydown(doc: Document) {
  const listeners = new Set<EventListenerOrEventListenerObject>();
  const add = doc.addEventListener,
    remove = doc.removeEventListener;
  doc.addEventListener = function (
    type: string,
    listener: EventListenerOrEventListenerObject,
    options?: boolean | AddEventListenerOptions
  ) {
    if (type === 'keydown') listeners.add(listener);
    add.call(doc, type, listener, options);
  };
  doc.removeEventListener = function (
    type: string,
    listener: EventListenerOrEventListenerObject,
    options?: boolean | EventListenerOptions
  ) {
    if (type === 'keydown') listeners.delete(listener);
    remove.call(doc, type, listener, options);
  };
  return {
    listeners,
    restore() {
      doc.addEventListener = add;
      doc.removeEventListener = remove;
    },
  };
}

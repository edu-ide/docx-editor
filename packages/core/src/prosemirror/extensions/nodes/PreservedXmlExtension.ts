/**
 * Preserved XML Extension — a run child the editor keeps as it was parsed
 *
 * A Word shape, chart, SmartArt, VML shape or embedded object the editor does
 * not model travels as an atom carrying its original XML, so saving writes it
 * back unchanged (see docx/preservedRunXml). `preview` holds the image-node
 * attributes of what it paints as — a line or rule, or a frame keeping the
 * shape's place — or null when it paints nothing.
 */

import { createNodeExtension } from '../create';

export const PreservedXmlExtension = createNodeExtension({
  name: 'preservedXml',
  schemaNodeName: 'preservedXml',
  nodeSpec: {
    inline: true,
    group: 'inline',
    atom: true,
    selectable: true,
    draggable: false,
    attrs: {
      /** The element exactly as parsed */
      xml: { default: '' },
      /** Image-node attributes of what it paints as, or null */
      preview: { default: null },
    },
    // No parseDOM: the XML can point at relationships of the document it came
    // from, so pasting must not carry it into another document.
    toDOM() {
      return ['span', { class: 'docx-preserved-xml', contenteditable: 'false' }];
    },
  },
});

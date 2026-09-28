/**
 * Write a parsed XML element back as XML.
 *
 * xml-js's js2xml escapes only `"` in attribute values: an `&` or `<` a value
 * held came back raw (invalid XML), and a `&#xA;` came back as a newline, which
 * an XML reader turns into a space — VML keeps base64 picture data that way.
 * Content kept as parsed must be written back as it was read, so attribute
 * values and text are both escaped here.
 */

import type { Element as XmlElement } from 'xml-js';

function escapeAttribute(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/"/g, '&quot;')
    .replace(/\t/g, '&#x9;')
    .replace(/\n/g, '&#xA;')
    .replace(/\r/g, '&#xD;');
}

function escapeText(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Write a node and its subtree as XML; an element without children self-closes. */
export function writeElement(node: XmlElement): string {
  switch (node.type) {
    case 'text':
      return escapeText(String(node.text ?? ''));
    case 'cdata':
      return `<![CDATA[${String(node.cdata ?? '')}]]>`;
    case 'element': {
      const attributes = Object.entries(node.attributes ?? {})
        .filter(([, value]) => value !== null && value !== undefined)
        .map(([name, value]) => ` ${name}="${escapeAttribute(String(value))}"`)
        .join('');
      const children = (node.elements ?? []).map(writeElement).join('');
      return children
        ? `<${node.name}${attributes}>${children}</${node.name}>`
        : `<${node.name}${attributes}/>`;
    }
    default:
      // parseXml keeps no comments, instructions or doctypes.
      return '';
  }
}

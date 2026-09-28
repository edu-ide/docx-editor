/**
 * Content kept as parsed is written back through elementToXml, so what it
 * writes must read back as the same values: xml-js's own writer left `&` and
 * `<` raw in attribute values and turned `&#xA;` into a newline an XML reader
 * reads as a space (VML base64 picture data is stored that way).
 */
import { describe, expect, test } from 'bun:test';
import { elementToXml, parseXml } from './xmlParser';
import type { XmlElement } from './xmlParser';

/** The document element of `xml`. */
function rootOf(xml: string): XmlElement {
  const root = parseXml(xml).elements?.find((node) => node.type === 'element');
  if (!root) throw new Error('no element');
  return root;
}

describe('elementToXml writes an element back as it was read', () => {
  test('attribute values keep entities, quotes and encoded whitespace', () => {
    const xml =
      '<v:shape xmlns:v="urn:v" title="a &amp; b &lt; c &quot;d&quot;" data="AB&#xA;CD&#x9;EF&#xD;"/>';
    expect(elementToXml(rootOf(xml))).toBe(
      '<v:shape xmlns:v="urn:v" title="a &amp; b &lt; c &quot;d&quot;" data="AB&#xA;CD&#x9;EF&#xD;"/>'
    );
    const reread = rootOf(elementToXml(rootOf(xml)));
    expect(reread.attributes).toEqual(rootOf(xml).attributes);
  });

  test('text keeps its entities and the whitespace between elements', () => {
    const xml =
      '<w:p xmlns:w="urn:w">\n  <w:r><w:t xml:space="preserve"> a &amp; b &lt; c &gt; d </w:t></w:r>\n</w:p>';
    expect(elementToXml(rootOf(xml))).toBe(xml);
  });

  test('an empty element self-closes and attribute order is kept', () => {
    expect(elementToXml(rootOf('<a:ext cy="0" cx="914400"></a:ext>'))).toBe(
      '<a:ext cy="0" cx="914400"/>'
    );
  });
});

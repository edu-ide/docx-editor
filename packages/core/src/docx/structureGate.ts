/**
 * Structure-preservation gate for DOCX repack.
 *
 * The repacker re-serializes word/document.xml from the editor's model. Elements
 * the editor does NOT model — text boxes (w:txbxContent), VML pictures (w:pict),
 * compatibility fallbacks (mc:AlternateContent) and embedded image blips
 * (a:blip) — must be passed through untouched. eigenpal v1.3.3 dropped ALL of
 * these from a real government form (-58% of document.xml) with zero warnings.
 *
 * This gate counts structural tags before vs after a repack and throws when a
 * *fragile* (unmodeled-passthrough) tag count drops — turning silent content
 * loss into a loud, blocking failure. Counts of legitimately-editable
 * structures (w:p, w:tbl, ...) are reported for context but do NOT trigger the
 * throw by default, so ordinary paragraph/table deletion in the editor is not a
 * false positive. Mirrors doc-mcp's Python `diff_docx_structure` gate.
 * @packageDocumentation
 */

/** Every structural tag we census (mirrors doc-mcp's diff_docx_structure set). */
export const STRUCTURAL_TAGS = [
  'w:p',
  'w:tbl',
  'w:tr',
  'w:tc',
  'w:drawing',
  'w:txbxContent',
  'w:pict',
  'mc:AlternateContent',
  'w:hyperlink',
  'a:blip',
  'w:sectPr',
] as const;

/**
 * Tags the editor passes through rather than models; a drop here ≈ silent loss,
 * not an intentional edit. These are the hard-gate triggers (the exact classes
 * the eigenpal -58% regression dropped).
 */
export const FRAGILE_TAGS = ['w:txbxContent', 'w:pict', 'mc:AlternateContent', 'a:blip'] as const;

export type StructuralTag = (typeof STRUCTURAL_TAGS)[number];

/**
 * Count occurrences of each structural start-tag in an XML string. Matches the
 * opening tag only (self-closing included) and uses a separator lookahead so a
 * prefix tag never over-counts a longer sibling: `w:p` ignores `w:pPr`,
 * `w:pict`, `w:proofErr`; closing tags (`</w:p>`) are not counted.
 */
export function countStructuralTags(xml: string): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const tag of STRUCTURAL_TAGS) {
    // `<tag` immediately followed by whitespace, '>', or '/' (self-closing).
    const re = new RegExp('<' + tag + '(?=[\\s/>])', 'g');
    counts[tag] = (xml.match(re) || []).length;
  }
  return counts;
}

export interface StructureDiff {
  tags: Record<string, { before: number; after: number; delta: number }>;
  /** Every structural tag whose count decreased. */
  lost: string[];
  /** The subset of `lost` that is unmodeled-passthrough (the hard-gate trigger). */
  fragileLost: string[];
  /** True when no fragile tag was lost. */
  ok: boolean;
}

export function diffStructuralTags(beforeXml: string, afterXml: string): StructureDiff {
  const before = countStructuralTags(beforeXml);
  const after = countStructuralTags(afterXml);
  const tags: StructureDiff['tags'] = {};
  const lost: string[] = [];
  for (const tag of STRUCTURAL_TAGS) {
    const delta = after[tag] - before[tag];
    tags[tag] = { before: before[tag], after: after[tag], delta };
    if (delta < 0) lost.push(tag);
  }
  const fragile = new Set<string>(FRAGILE_TAGS);
  const fragileLost = lost.filter((t) => fragile.has(t));
  return { tags, lost, fragileLost, ok: fragileLost.length === 0 };
}

/**
 * Throw if a repack would silently drop unmodeled-passthrough structures.
 *
 * @param strict - also throw on the loss of editable structures (w:p/w:tbl/...).
 *   Use it for an explicit "save edited bytes" path; leave it off for the live
 *   collaborative editor, where deleting a paragraph is a legitimate edit.
 * @returns the full structural diff (for logging) when no throw condition is met.
 */
export function assertNoStructuralLoss(
  beforeXml: string,
  afterXml: string,
  opts: { strict?: boolean } = {}
): StructureDiff {
  const diff = diffStructuralTags(beforeXml, afterXml);
  const trigger = opts.strict ? diff.lost : diff.fragileLost;
  if (trigger.length > 0) {
    const detail = trigger.map((t) => `${t} ${diff.tags[t].before}->${diff.tags[t].after}`).join(', ');
    throw new Error(
      `Repack blocked: structural content would be silently lost (${detail}). ` +
        'The edited model is missing elements present in the original document.xml; ' +
        'saving would corrupt the file. This is the eigenpal txbxContent/pict/' +
        'AlternateContent loss class — fix the serializer, or pass skipStructureGate to override.'
    );
  }
  return diff;
}

---
'@eigenpal/docx-editor-core': patch
---

An empty section-break paragraph that follows content in its own section ignores its own page break before, from direct formatting or its style, and no longer opens a blank page before the next section. Saving keeps each page break before as the document had it: a style's page break is no longer written into the paragraph as direct formatting, a direct `w:val="0"` that switches a style's break off is kept, and a page break that opens a paragraph stays a page-break run instead of becoming a `w:pageBreakBefore` property. Backport of upstream 9afb832b (#983).

---
'@eigenpal/docx-editor-core': patch
'@eigenpal/docx-editor-vue': patch
---

Backport OSS layout corrections while retaining the existing editor and collaboration APIs: split oversized cantSplit table rows at whole-line boundaries, select East Asian font slots for mixed-script painted runs without changing document positions, and parse XML Schema boolean anchor attributes. Make word-boundary traversal linear for long space-heavy paragraphs.

Allow Vue hosts to configure the same per-peer comment ID allocator policy as React.

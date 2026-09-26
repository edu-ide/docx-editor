---
'@eigenpal/docx-editor-core': patch
---

Resolve a paragraph mark's run properties above the paragraph style: the character style its `w:rStyle` names, then its direct properties. Any mark property (such as bold) no longer brings the document default font size back over the paragraph style's for empty paragraphs, empty list markers and newly typed text, and an `w:rStyle` that names a paragraph style is ignored. Backport of upstream 2eea4deb (#987).

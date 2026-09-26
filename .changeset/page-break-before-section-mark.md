---
'@eigenpal/docx-editor-core': patch
---

A manual page break followed by an empty section-break paragraph no longer adds a blank page before a section that starts on a new page: the break and the empty mark stay on the page the break closes. Backport of upstream d6c75d2c (#981).

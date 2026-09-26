---
'@eigenpal/docx-editor-core': patch
---

Start a hanging numbered paragraph's first line at an authored tab stop that lies between the number and the text indent, so the line wraps with the width it has, and start it right after the number for a `space` or `nothing` suffix. Backport of upstream bf776f2d (#979).

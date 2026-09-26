---
'@eigenpal/docx-editor-core': patch
---

End a keep-with-next group at a forced break: a heading whose next paragraph starts a new page, or that is followed by a page or column break, stays on its page instead of moving to a page of its own, and the break discards its trailing spacing. Backport of upstream ab460dc9 (#985).

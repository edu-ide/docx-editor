---
'@eigenpal/docx-editor-core': patch
---

Keep no-break spaces (U+00A0, U+2007, U+202F) and word joiners (U+2060, U+FEFF) with their neighbours across formatting runs, and keep a format-only split inside a word from wrapping at the run boundary: text split into runs now wraps exactly like the same text in one run. Backport of upstream ae1afe03 (#976).

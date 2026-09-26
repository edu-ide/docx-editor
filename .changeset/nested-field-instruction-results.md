---
'@eigenpal/docx-editor-core': patch
---

A field nested inside another field's instruction (such as `IF { STYLEREF Heading } <> "x" ...`) no longer displays its cached result as text of its own, and a nested complex field no longer breaks up the outer field. The nested fields stay in the outer field code, and saving writes them back inside the instruction. Backport of upstream 390c1772 (#986).

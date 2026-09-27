---
'@eigenpal/docx-editor-core': patch
---

Paint straight lines, connectors and outlined horizontal rules drawn as `wps:wsp` shapes: each now draws its stroke instead of an empty picture box, and a vertical or horizontal line takes its stroke width on the axis it has no size on instead of 100px, so it no longer pushes its line of text down. The line geometry is on `Image.vectorShape` and the image node's `vectorShape` attribute; saving writes these drawings as before. Backport of upstream d04902a9 (#972).

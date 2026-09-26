---
'@eigenpal/docx-editor-core': patch
---

Start an odd-page or even-page section on a page number of that parity: when the page after the break has the other parity, layout now inserts one blank page (`Page.parityBlank`), sized like the section it precedes and painted without header, footer, watermark or page border. The blank page takes a page number and counts toward the page total. Backport of upstream 6794f4d3 (#978).

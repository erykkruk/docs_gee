# TODO - docs_gee

## High Priority

- [x] **Arabic/RTL Support (DOCX)** - shipped in 1.5.0: `DocxParagraph.rtl` and `DocxRun.rtl` generate `<w:bidi/>` and `<w:rtl/>`, and round-trip through `DocxReader`.
- [ ] **Custom Styles (DOCX)** - Allow users to define custom paragraph and character styles

---

## Normal Priority

- [x] **Font Size per Run** - shipped in 1.5.0 as `DocxRun.fontSize`, honoured by both generators.
- [x] **Cell Padding for Tables** - shipped in 1.5.0 as `DocxCellPadding` on `DocxTable.cellPadding` and `DocxTableCell.padding`.
- [x] **Images Support (DOCX)** - shipped in 1.5.0 as `DocxImage` plus `DocxDocument.addImage`.

---

## Low Priority

- [ ] **PDF Font Embedding** - Embed TTF/OTF fonts for extended character support (CJK, Hindi, full Polish). Very complex - requires font parsing and subsetting.
- [ ] **PDF RTL Support** - Right-to-left text in PDF. Very complex - requires text shaping library.
- [x] **Headers and Footers (DOCX)** - shipped in 1.5.0 as `DocxHeaderFooter` on `DocxDocument.header` / `.footer`.
- [x] **Page Numbers (DOCX)** - shipped in 1.5.0 as `DocxRun.pageNumber()` / `DocxRun.pageCount()` and `DocxHeaderFooter.pageNumber()`.
- [ ] **Footnotes/Endnotes (DOCX)** - Proper footnote support with automatic numbering and references
- [ ] **PDF Hyperlinks** - Add clickable external links in PDF output

---

## Completed (v1.1.2)

- [x] Fix bullet points in PDF (encoding issue showing "Â•")
- [x] Fix list detection for dash/alpha/roman styles in DOCX
- [x] Add extended character support in PDF (German, French, typography symbols)
- [x] Add Polish character support in PDF (Ó/ó native, others fallback to ASCII)
- [x] Add `<w:ilvl>` to list styles in DOCX for consistent formatting

---

## Won't Fix / Limitations

- **Full Polish in PDF** - ą, ę, ć, ź, ż, ń, ł, ś not in WinAnsi. Workaround: ASCII fallback.
- **Emoji in PDF** - Base14 fonts lack emoji glyphs. Workaround: use DOCX.
- **CJK in PDF** - Chinese/Japanese/Korean require font embedding. Workaround: use DOCX.
- **Hindi in PDF** - Requires complex text shaping. Workaround: use DOCX.

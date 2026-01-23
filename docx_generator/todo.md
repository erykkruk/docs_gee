# TODO - docs_gee

## High Priority

- [ ] **Arabic/RTL Support (DOCX)** - Add right-to-left text direction for Arabic, Hebrew, Persian. Add `rtl` property to paragraphs and runs, generate `<w:bidi/>` and `<w:rtl/>` XML elements.
- [ ] **Custom Styles (DOCX)** - Allow users to define custom paragraph and character styles

---

## Normal Priority

- [ ] **Font Size per Run** - Allow specifying font size at the run level (not just document level) for both DOCX and PDF
- [ ] **Cell Padding for Tables** - Add configurable padding/margins for table cells in DOCX and PDF
- [ ] **Images Support (DOCX)** - Add ability to insert PNG/JPEG images into documents

---

## Low Priority

- [ ] **PDF Font Embedding** - Embed TTF/OTF fonts for extended character support (CJK, Hindi, full Polish). Very complex - requires font parsing and subsetting.
- [ ] **PDF RTL Support** - Right-to-left text in PDF. Very complex - requires text shaping library.
- [ ] **Headers and Footers (DOCX)** - Add page headers and footers with custom content
- [ ] **Page Numbers (DOCX)** - Add automatic page numbering (depends on headers/footers)
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

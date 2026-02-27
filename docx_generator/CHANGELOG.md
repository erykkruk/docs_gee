# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.2.2] - 2026-02-27

### Fixed

#### PDF Generator
- **Accurate text width estimation** - Replaced rough character-width approximation (`0.5 * fontSize`) with Adobe Helvetica AFM glyph-width table for precise text layout
  - Helvetica variants use per-character widths from the standard Adobe Font Metrics
  - Courier (monospace) uses correct `0.6 * fontSize` per glyph
  - Includes width data for Latin-1 supplement and Polish characters
- **Underline & strikethrough in table cells** - Table cell text segments now render underline and strikethrough decorations correctly (previously only worked outside tables)
- **Per-segment width calculation** - Line width is now computed per text segment (respecting each segment's font) instead of joining all text and measuring as one

## [1.2.1] - 2026-02-23

### Fixed

#### Page Breaks
- **Windows Word compatibility** - Page breaks now work correctly in Microsoft Word on Windows
  - Replaced `<w:pageBreakBefore/>` paragraph property with explicit `<w:br w:type="page"/>` run element
  - Applies to regular paragraphs, table cell paragraphs, and TOC page breaks
  - No API changes - `pageBreakBefore: true` works exactly the same way
- **Unit tests** - Added `test/page_break_test.dart` with 11 tests covering page break XML generation, DOCX and PDF output

## [1.2.0] - 2026-02-19

### Added

#### Per-Cell Border Control
- **Cell-level borders** - Override table-level borders on individual cells
  - `DocxCellBorders` class with `top`, `bottom`, `left`, `right` borders
  - `CellBorders` type alias for format-agnostic code
  - Convenience constructors: `.all()`, `.none()`, `.bottom()`
  - Custom border per side with `color`, `size`, and `style`
  - Cell borders override table-level borders when set
  - Works in both DOCX (`<w:tcBorders>`) and PDF output
- **`borders` parameter on `DocxTableCell`** - Both constructor and `.text()` factory accept `DocxCellBorders?`
- **Unit tests** - Added `test/cell_borders_test.dart` with 26 tests covering model, XML generation, and PDF/DOCX output

## [1.1.2] - 2025-01-23

### Fixed

#### PDF Generator
- **Bullet points display correctly** - Fixed encoding issue where bullet character (•) was displaying as "Â•"
  - Rewrote `_escapePdfString` to properly convert Unicode characters to WinAnsi octal escapes
- **Extended character support** - Added proper encoding for:
  - Typography: bullet (•), en-dash (–), em-dash (—), smart quotes (' ' " "), ellipsis (…), euro (€), trademark (™), copyright (©), registered (®)
  - German: Ä, Ö, Ü, ä, ö, ü, ß
  - French: À, Â, Ç, È, É, Ê, Ë, Î, Ï, Ô, Ù, Û, à, â, ç, è, é, ê, ë, î, ï, ô, ù, û
  - Polish: Ó, ó (native support), other Polish characters (ą, ę, ć, ź, ż, ń, ł, ś) fall back to base ASCII equivalents due to WinAnsi limitations

#### DOCX Generator
- **List detection fix** - Fixed issue where `listDash`, `listNumberAlpha`, and `listNumberRoman` styles didn't trigger inclusion of `numbering.xml`, causing lists to display incorrectly
- **List level specification** - Added explicit `<w:ilvl w:val="0"/>` to all list styles in `styles.xml` for consistent formatting

## [1.1.0] - 2025-01-21

### Added

#### Line Breaks (Soft Return) - DOCX only
- **Automatic `\n` conversion** - Newline characters in text are now converted to `<w:br/>` line breaks
  - `DocxParagraph.text('Line 1\nLine 2\nLine 3')`
- **Explicit line break runs** - New `DocxRun.lineBreak()` constructor for manual line breaks
  - Allows different formatting before and after line break
  - `DocxParagraph(runs: [DocxRun('Bold', bold: true), DocxRun.lineBreak(), DocxRun('Normal')])`
- Line breaks create soft returns within a paragraph (like Shift+Enter in Word)

## [1.0.1] - 2024-12-21

### Added

#### Hyperlinks
- **External links** - Link text to external URLs
  - `DocxRun('text', hyperlink: 'https://example.com')`
  - Automatic blue color and underline styling

#### Internal Links (Bookmarks)
- **Bookmarks** - Create named anchors in document
  - `DocxParagraph.heading('Title', level: 1, bookmarkName: 'section1')`
- **Bookmark references** - Link to internal bookmarks
  - `DocxRun('Go to section', bookmarkRef: 'section1')`

#### Table of Contents
- **Automatic TOC generation**
  - `DocxDocument(includeTableOfContents: true)`
  - Configurable title (`tocTitle`)
  - Configurable heading depth (`tocMaxLevel: 1-4`)
  - Word auto-updates TOC on document open

### Fixed
- Repository URL in pubspec.yaml

## [1.0.0] - 2024-12-19

### Added

#### Document Generation
- **DocxGenerator** - Generates Microsoft Word DOCX files
  - Pure Dart implementation using OOXML standard
  - Configurable default font name and size
  - Compatible with Word 2007+, Google Docs, LibreOffice
- **PdfGenerator** - Generates PDF documents
  - Pure Dart implementation (PDF 1.4 standard)
  - No native dependencies
  - Cross-platform support including Web

#### Document Model
- **Document** - Container for document content
  - Title and author metadata
  - Mixed content support (paragraphs and tables)
  - Creation and modification timestamps

#### Text Formatting
- **TextRun** - Rich text formatting
  - Bold, italic, underline, strikethrough
  - Text color (hex format)
  - Background/highlight color
- **Text alignment** - Left, center, right, justify

#### Paragraph Styles
- Headings (H1-H4)
- Subtitle
- Caption
- Block quote
- Code block (monospace)
- Footnote

#### Lists
- Bullet lists (•)
- Dash lists (–)
- Numbered lists (1, 2, 3)
- Alphabetic lists (a, b, c)
- Roman numeral lists (I, II, III)
- Nested lists up to 9 levels deep

#### Tables
- Basic table support with rows and cells
- Cell background colors
- Cell text alignment
- Table borders (all, none, outside only)
- Border styles (single, double, dashed, dotted)

#### Other Features
- Page breaks
- Emoji support (DOCX only - Word handles natively)
- Cross-platform: iOS, Android, Web, macOS, Windows, Linux

### Documentation
- Complete API reference
- Usage examples for all features
- Platform-specific guides (Web, Mobile)

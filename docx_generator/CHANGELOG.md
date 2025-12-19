# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2024-12-19

### Added

- Initial release of docs_gee library
- **DocxGenerator** - Main class for generating DOCX files
  - Configurable default font name and size
  - Generates valid OOXML-compliant documents
- **DocxDocument** - Document container
  - Support for title and author metadata
  - Methods to add single or multiple paragraphs
- **DocxParagraph** - Paragraph model with factory constructors
  - `DocxParagraph.text()` - Simple text paragraph
  - `DocxParagraph.heading()` - Heading levels 1-3
  - `DocxParagraph.bulletItem()` - Bullet list items
  - `DocxParagraph.numberedItem()` - Numbered list items
- **DocxRun** - Text formatting model
  - Bold, italic, underline, strikethrough support
  - `copyWith()` method for easy modification
- **DocxAlignment** - Text alignment enum
  - Left, center, right, justify options
- **DocxParagraphStyle** - Paragraph style enum
  - Normal, heading1-3, listBullet, listNumber
- **Page breaks** - Control pagination with `pageBreakBefore`
- **Cross-platform support** - Works on iOS, Android, Web, macOS, Windows, Linux
- Complete documentation with API reference
- Example application demonstrating all features

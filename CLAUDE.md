# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is `docx_generator`, a pure Dart library for generating DOCX (Word) files without external dependencies beyond the `archive` package for ZIP creation.

## Commands

```bash
# Get dependencies
cd docx_generator && dart pub get

# Run the example
cd docx_generator && dart run example/example.dart

# Analyze code
cd docx_generator && dart analyze

# Format code
cd docx_generator && dart format .
```

## Architecture

The library generates DOCX files by assembling XML parts into a ZIP archive. DOCX is essentially a ZIP file containing XML documents.

### Core Components

**Models** (`lib/src/models/`):
- `DocxDocument` - Container for paragraphs with optional title/author metadata
- `DocxParagraph` - Paragraph with runs, style, alignment, and page break support. Factory constructors: `.text()`, `.heading()`, `.bulletItem()`, `.numberedItem()`
- `DocxRun` - Text segment with formatting (bold, italic, underline, strikethrough)
- `DocxAlignment` / `DocxParagraphStyle` - Enums for alignment and paragraph styles

**XML Parts** (`lib/src/xml_parts/`):
- `DocumentXml` - Main content (word/document.xml)
- `StylesXml` - Document styles with configurable font/size
- `NumberingXml` - List numbering definitions (bullet/numbered)
- `ContentTypesXml` - MIME type declarations
- `RelsXml` - Relationship files linking document parts
- `XmlUtils` - XML escaping and namespace constants

**Generator** (`lib/src/docx_generator.dart`):
- `DocxGenerator` - Assembles all XML parts into a ZIP archive, returns `Uint8List`
- Configurable default font name and size (in half-points: 24 = 12pt)

### Data Flow

1. User creates `DocxDocument` and adds `DocxParagraph` objects
2. `DocxGenerator.generate()` creates XML for each required part
3. XML files are UTF-8 encoded and added to an `Archive`
4. Archive is ZIP-encoded and returned as bytes

### DOCX Structure

Generated files contain:
- `[Content_Types].xml` - MIME declarations
- `_rels/.rels` - Root relationships
- `word/document.xml` - Main content
- `word/styles.xml` - Style definitions
- `word/_rels/document.xml.rels` - Document relationships
- `word/numbering.xml` - Only included when document contains lists

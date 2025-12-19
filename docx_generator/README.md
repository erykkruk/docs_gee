# docs_gee

A pure Dart library for generating **DOCX** and **PDF** documents. Works on all platforms: iOS, Android, Web, macOS, Windows, and Linux.

## Features

- **Pure Dart** - No native dependencies, works everywhere Dart runs
- **Dual format** - Generate both DOCX and PDF from the same document model
- **Cross-platform** - iOS, Android, Web, Desktop (macOS, Windows, Linux)
- **Rich text formatting** - Bold, italic, underline, strikethrough, colors
- **Paragraph styles** - Headings (H1-H4), subtitle, caption, quote, code block, footnote
- **Text alignment** - Left, center, right, justify
- **Lists** - Bullet, dash, numbered (1,2,3), alphabetic (a,b,c), roman (I,II,III)
- **Nested lists** - Up to 9 levels of nesting
- **Tables** - Rows, cells, borders, cell background colors
- **Emoji support** - Works in DOCX (Word handles natively)
- **Page breaks** - Control document pagination
- **Text colors** - Foreground and background/highlight colors
- **Configurable fonts** - Set default font and size
- **Document metadata** - Title, author, creation date
- **Lightweight** - Minimal dependencies (`archive` for ZIP)

## Installation

Add to your `pubspec.yaml`:

```yaml
dependencies:
  docs_gee: ^1.0.0
```

Or install via command line:

```bash
dart pub add docs_gee
# or for Flutter projects:
flutter pub add docs_gee
```

## Quick Start

```dart
import 'dart:io';
import 'package:docs_gee/docs_gee.dart';

void main() {
  // Create document
  final doc = DocxDocument(
    title: 'My Document',
    author: 'John Doe',
  );

  // Add content
  doc.addParagraph(DocxParagraph.heading('Hello World', level: 1));
  doc.addParagraph(DocxParagraph.text('This is a simple paragraph.'));
  doc.addParagraph(DocxParagraph.bulletItem('First item'));
  doc.addParagraph(DocxParagraph.bulletItem('Second item'));

  // Generate DOCX
  final docxBytes = DocxGenerator().generate(doc);
  File('output.docx').writeAsBytesSync(docxBytes);

  // Generate PDF (same document!)
  final pdfBytes = PdfGenerator().generate(doc);
  File('output.pdf').writeAsBytesSync(pdfBytes);
}
```

### Run the Full Demo

To see all features in action, run the comprehensive example:

```bash
cd docx_generator
dart run example/all_features_example.dart
```

This generates `all_features.docx` and `all_features.pdf` showcasing every feature.

## Usage Examples

### Text Formatting

```dart
doc.addParagraph(DocxParagraph(
  runs: [
    const DocxRun('Normal text, '),
    const DocxRun('bold text, ', bold: true),
    const DocxRun('italic text, ', italic: true),
    const DocxRun('underlined text, ', underline: true),
    const DocxRun('strikethrough text.', strikethrough: true),
  ],
));
```

### Headings

```dart
doc.addParagraph(DocxParagraph.heading('Main Title', level: 1));
doc.addParagraph(DocxParagraph.heading('Section', level: 2));
doc.addParagraph(DocxParagraph.heading('Subsection', level: 3));
doc.addParagraph(DocxParagraph.heading('Minor Section', level: 4));
```

### Semantic Styles

```dart
doc.addParagraph(DocxParagraph.subtitle('Document subtitle'));
doc.addParagraph(DocxParagraph.caption('Figure 1: Example caption'));
doc.addParagraph(DocxParagraph.quote('This is a blockquote...'));
doc.addParagraph(DocxParagraph.codeBlock('const x = 42;\nconsole.log(x);'));
doc.addParagraph(DocxParagraph.footnote('1. Reference note here.'));
```

### Text Colors

```dart
doc.addParagraph(DocxParagraph(
  runs: [
    const DocxRun('Red text ', color: 'FF0000'),
    const DocxRun('with yellow highlight', backgroundColor: 'FFFF00'),
  ],
));
```

### Text Alignment

```dart
doc.addParagraph(DocxParagraph.text('Left aligned'));
doc.addParagraph(DocxParagraph.text('Centered', alignment: DocxAlignment.center));
doc.addParagraph(DocxParagraph.text('Right aligned', alignment: DocxAlignment.right));
doc.addParagraph(DocxParagraph.text('Justified text...', alignment: DocxAlignment.justify));
```

### Bullet Lists

```dart
doc.addParagraph(DocxParagraph.bulletItem('First item'));
doc.addParagraph(DocxParagraph.bulletItem('Second item'));
doc.addParagraph(DocxParagraph.bulletItem('Third item'));
```

### Dash Lists

```dart
doc.addParagraph(DocxParagraph.dashItem('First item'));
doc.addParagraph(DocxParagraph.dashItem('Second item'));
doc.addParagraph(DocxParagraph.dashItem('Third item'));
```

### Numbered Lists

```dart
// Numeric (1, 2, 3...)
doc.addParagraph(DocxParagraph.numberedItem('Step one'));
doc.addParagraph(DocxParagraph.numberedItem('Step two'));

// Alphabetic (a, b, c...)
doc.addParagraph(DocxParagraph.alphaItem('Item a'));
doc.addParagraph(DocxParagraph.alphaItem('Item b'));

// Roman numerals (I, II, III...)
doc.addParagraph(DocxParagraph.romanItem('Section I'));
doc.addParagraph(DocxParagraph.romanItem('Section II'));
```

### Nested Lists

```dart
doc.addParagraph(DocxParagraph.bulletItem('Top level'));
doc.addParagraph(DocxParagraph.bulletItem('Nested item', indentLevel: 1));
doc.addParagraph(DocxParagraph.bulletItem('Deep nested', indentLevel: 2));
doc.addParagraph(DocxParagraph.bulletItem('Back to top level'));
```

### Tables

```dart
doc.addTable(DocxTable(
  borders: const DocxTableBorders.all(),
  rows: [
    DocxTableRow(cells: [
      DocxTableCell.text('Name', backgroundColor: 'E0E0E0'),
      DocxTableCell.text('Age', backgroundColor: 'E0E0E0'),
    ]),
    DocxTableRow(cells: [
      DocxTableCell.text('Alice'),
      DocxTableCell.text('30', alignment: DocxAlignment.right),
    ]),
    DocxTableRow(cells: [
      DocxTableCell.text('Bob'),
      DocxTableCell.text('25', alignment: DocxAlignment.right),
    ]),
  ],
));
```

### Emoji (DOCX only)

```dart
// Emoji works in DOCX - Word handles it natively
doc.addParagraph(DocxParagraph.text('Hello World! 👋🌍'));
doc.addParagraph(DocxParagraph.text('Status: ✅ Complete'));
```

### Page Breaks

```dart
doc.addParagraph(DocxParagraph.text('Content on first page'));
doc.addParagraph(DocxParagraph.heading(
  'New Chapter',
  level: 1,
  pageBreakBefore: true,
));
```

### Custom Fonts

```dart
final generator = DocxGenerator(
  fontName: 'Arial',
  fontSize: 28, // 14pt (size is in half-points)
);
```

## Platform-Specific Notes

### Web

On web, you cannot use `dart:io`. Use the generated bytes with web APIs:

```dart
import 'dart:html' as html;

void downloadDocx(Uint8List bytes) {
  final blob = html.Blob([bytes], 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  final url = html.Url.createObjectUrlFromBlob(blob);
  html.AnchorElement(href: url)
    ..setAttribute('download', 'document.docx')
    ..click();
  html.Url.revokeObjectUrl(url);
}
```

### Mobile (iOS/Android)

Save to app documents directory or share:

```dart
import 'package:path_provider/path_provider.dart';

Future<File> saveDocx(Uint8List bytes) async {
  final dir = await getApplicationDocumentsDirectory();
  final file = File('${dir.path}/document.docx');
  await file.writeAsBytes(bytes);
  return file;
}
```

### Desktop (macOS/Windows/Linux)

Use standard file operations with `dart:io`.

## How It Works

DOCX files are ZIP archives containing XML documents. This library:

1. Creates XML for each required part (document content, styles, relationships)
2. Packages everything into a ZIP archive using the `archive` package
3. Returns the bytes ready for saving or transmission

The generated files are compatible with:
- Microsoft Word (2007+)
- Google Docs
- LibreOffice Writer
- Apple Pages
- And other OOXML-compatible applications

## API Reference

See the full [API documentation](doc/api.md) for detailed reference.

### Main Classes

| Class | Description |
|-------|-------------|
| `DocxGenerator` | Generates DOCX bytes from a document |
| `DocxDocument` | Container for document content and metadata |
| `DocxParagraph` | Paragraph with text, style, and alignment |
| `DocxRun` | Text segment with formatting properties |

### Enums

| Enum | Values |
|------|--------|
| `DocxAlignment` | `left`, `center`, `right`, `justify` |
| `DocxParagraphStyle` | `normal`, `heading1`, `heading2`, `heading3`, `listBullet`, `listNumber` |

## Requirements

- Dart SDK: >=3.0.0
- Flutter: >=3.0.0 (if using with Flutter)

## Feature Status

Tabela pokazuje aktualny stan implementacji funkcji dla DOCX i PDF.

### 1. Text – Inline styles (TextRun / Span)

| Funkcja | DOCX | PDF |
|---------|:----:|:---:|
| **bold** | ✅ | ✅ |
| **italic** | ✅ | ✅ |
| **underline** | ✅ | ✅ |
| **strikethrough** | ✅ | ✅ |
| superscript | ❌ | ❌ |
| subscript | ❌ | ❌ |
| **font family** | ✅ | ✅ |
| **font size (pt)** | ✅ | ✅ |
| **text color** | ✅ | ✅ |
| **background / highlight color** | ✅ | ❌ |
| letter spacing | ❌ | ❌ |
| word spacing | ❌ | ❌ |
| text shadow | ❌ | ❌ |
| all caps / small caps | ❌ | ❌ |
| **monospace (via code block)** | ✅ | ✅ |

### 2. Paragraph & Block structure

| Funkcja | DOCX | PDF |
|---------|:----:|:---:|
| **new paragraph** | ✅ | ✅ |
| **empty paragraph** | ✅ | ✅ |
| paragraph spacing (before / after) | ❌ | ❌ |
| line height | ❌ | ❌ |
| **text alignment: left** | ✅ | ✅ |
| **text alignment: right** | ✅ | ✅ |
| **text alignment: center** | ✅ | ✅ |
| **text alignment: justify** | ✅ | ✅ |
| first line indent | ❌ | ❌ |
| left indent | ❌ | ❌ |
| right indent | ❌ | ❌ |

### 3. Semantic text styles

| Funkcja | DOCX | PDF |
|---------|:----:|:---:|
| **normal text / paragraph** | ✅ | ✅ |
| **H1** | ✅ | ✅ |
| **H2** | ✅ | ✅ |
| **H3** | ✅ | ✅ |
| **H4** | ✅ | ✅ |
| **subtitle** | ✅ | ✅ |
| **caption** | ✅ | ✅ |
| **quote / blockquote** | ✅ | ✅ |
| **code block** | ✅ | ✅ |
| **footnote text** | ✅ | ✅ |
| small print / disclaimer | ❌ | ❌ |

### 4. Lists

| Funkcja | DOCX | PDF |
|---------|:----:|:---:|
| **unordered list (bullet •)** | ✅ | ✅ |
| **unordered list (dash -)** | ✅ | ✅ |
| unordered list (custom symbol) | ❌ | ❌ |
| **ordered list (1, 2, 3)** | ✅ | ✅ |
| **ordered list (a, b, c)** | ✅ | ✅ |
| **ordered list (I, II, III)** | ✅ | ✅ |
| **nested lists** | ✅ | ✅ |
| list item spacing | ❌ | ❌ |
| start index | ❌ | ❌ |

### 5. Tables

| Funkcja | DOCX | PDF |
|---------|:----:|:---:|
| **table** | ✅ | ✅ |
| **row** | ✅ | ✅ |
| **cell** | ✅ | ✅ |
| rowspan | ❌ | ❌ |
| colspan | ❌ | ❌ |
| table width (auto) | ✅ | ✅ |
| column widths (equal) | ✅ | ✅ |
| cell padding | ✅ | ✅ |
| **cell alignment** | ✅ | ✅ |
| vertical alignment | ❌ | ❌ |
| **borders** | ✅ | ✅ |
| **background color per cell** | ✅ | ✅ |
| header row | ❌ | ❌ |
| repeat header row | ❌ | ❌ |

### 6. Images & Media

| Funkcja | DOCX | PDF |
|---------|:----:|:---:|
| image from asset | ❌ | ❌ |
| image from file | ❌ | ❌ |
| image from bytes | ❌ | ❌ |
| image from URL | ❌ | ❌ |
| width / height | ❌ | ❌ |
| keep aspect ratio | ❌ | ❌ |
| alignment | ❌ | ❌ |
| inline image | ❌ | ❌ |
| block image | ❌ | ❌ |
| caption | ❌ | ❌ |
| image padding | ❌ | ❌ |
| image border | ❌ | ❌ |

### 7. Page & Layout

| Funkcja | DOCX | PDF |
|---------|:----:|:---:|
| page size (A4) | ❌ | ✅ |
| page size (Letter) | ❌ | ✅ |
| page size (custom) | ❌ | ✅ |
| orientation (portrait) | ❌ | ✅ |
| orientation (landscape) | ❌ | ❌ |
| **margins** | ❌ | ✅ |
| header | ❌ | ❌ |
| footer | ❌ | ❌ |
| different first page | ❌ | ❌ |
| different odd/even pages | ❌ | ❌ |
| page number | ❌ | ❌ |
| page total count | ❌ | ❌ |
| dynamic text (date, title, author) | ❌ | ❌ |

### 8. Sections & Flow control

| Funkcja | DOCX | PDF |
|---------|:----:|:---:|
| section break | ❌ | ❌ |
| **page break** | ✅ | ✅ |
| keep together | ❌ | ❌ |
| keep with next | ❌ | ❌ |
| force page break before | ✅ | ✅ |
| columns (1–3) | ❌ | ❌ |

### 9. Document metadata

| Funkcja | DOCX | PDF |
|---------|:----:|:---:|
| **title** | ✅ | ✅ |
| **author** | ✅ | ✅ |
| subject | ❌ | ❌ |
| keywords | ❌ | ❌ |
| **creation date** | ❌ | ✅ |
| **modification date** | ❌ | ✅ |
| language | ❌ | ❌ |
| reading direction (LTR / RTL) | ❌ | ❌ |

### 10. Interactivity (PDF only)

| Funkcja | PDF |
|---------|:---:|
| clickable links | ❌ |
| mailto links | ❌ |
| internal anchors | ❌ |
| table of contents (TOC) | ❌ |
| bookmarks / outline | ❌ |
| form fields | ❌ |

### 11. Internationalization & Typography

| Funkcja | DOCX | PDF |
|---------|:----:|:---:|
| UTF-8 support | ✅ | ⚠️ |
| **emoji support** | ✅ | ❌ |
| RTL languages | ❌ | ❌ |
| hyphenation | ❌ | ❌ |
| widows & orphans control | ❌ | ❌ |
| ligatures | ❌ | ❌ |

**Legenda:** ✅ Zaimplementowane | ❌ Do zrobienia | ⚠️ Częściowo (tylko ASCII/WinAnsi)

---

## License

MIT License - see [LICENSE](LICENSE) for details.

## Contributing

Contributions are welcome! Please feel free to submit issues and pull requests.

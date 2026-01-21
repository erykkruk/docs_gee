# docs_gee

[![pub package](https://img.shields.io/pub/v/docs_gee.svg)](https://pub.dev/packages/docs_gee)
[![likes](https://img.shields.io/pub/likes/docs_gee)](https://pub.dev/packages/docs_gee/score)
[![popularity](https://img.shields.io/pub/popularity/docs_gee)](https://pub.dev/packages/docs_gee/score)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

A **pure Dart** library for generating **Microsoft Word (DOCX)** and **PDF** documents. Create professional documents programmatically with rich formatting, tables, lists, and more - all from a single document model. Created and supported by [Codigee](https://umami.team.codigee.com/q/FZ9PQYyYN).

## Why docs_gee?

- **Pure Dart** - No native dependencies, works everywhere Dart runs
- **Dual Format** - Generate both DOCX and PDF from the same document model
- **Cross-Platform** - iOS, Android, Web, macOS, Windows, Linux
- **Lightweight** - Only one dependency (`archive` for ZIP)
- **Simple API** - Intuitive document builder pattern

## Features

| Feature | DOCX | PDF |
|---------|:----:|:---:|
| Text formatting (bold, italic, underline, strikethrough) | ✅ | ✅ |
| Text colors & highlighting | ✅ | ✅ |
| Headings (H1-H4) | ✅ | ✅ |
| Paragraph styles (subtitle, caption, quote, code, footnote) | ✅ | ✅ |
| Text alignment (left, center, right, justify) | ✅ | ✅ |
| Bullet & numbered lists | ✅ | ✅ |
| Nested lists (up to 9 levels) | ✅ | ✅ |
| Tables with borders & colors | ✅ | ✅ |
| Page breaks | ✅ | ✅ |
| Line breaks (soft return) | ✅ | - |
| Hyperlinks (external URLs) | ✅ | - |
| Internal links (bookmarks) | ✅ | - |
| Table of Contents | ✅ | - |
| Document metadata | ✅ | ✅ |
| Custom fonts | ✅ | ✅ |
| Emoji support | ✅ | - |

## Installation

```yaml
dependencies:
  docs_gee: ^1.1.0
```

```bash
dart pub add docs_gee
# or
flutter pub add docs_gee
```

## Quick Start

```dart
import 'dart:io';
import 'package:docs_gee/docs_gee.dart';

void main() {
  // Create document
  final doc = Document(title: 'My Report', author: 'John Doe');

  // Add content
  doc.addParagraph(Paragraph.heading('Quarterly Report', level: 1));
  doc.addParagraph(Paragraph.text('This report summarizes Q4 performance.'));

  // Add a table
  doc.addTable(Table(
    rows: [
      TableRow(cells: [
        TableCell.text('Metric', backgroundColor: 'E0E0E0'),
        TableCell.text('Value', backgroundColor: 'E0E0E0'),
      ]),
      TableRow(cells: [
        TableCell.text('Revenue'),
        TableCell.text('\$1.2M', alignment: Alignment.right),
      ]),
    ],
  ));

  // Generate both formats
  File('report.docx').writeAsBytesSync(DocxGenerator().generate(doc));
  File('report.pdf').writeAsBytesSync(PdfGenerator().generate(doc));
}
```

## Usage Examples

### Rich Text Formatting

```dart
doc.addParagraph(Paragraph(
  runs: [
    TextRun('Normal, '),
    TextRun('bold, ', bold: true),
    TextRun('italic, ', italic: true),
    TextRun('colored', color: 'FF0000'),
  ],
));
```

### Lists

```dart
// Bullet list
doc.addParagraph(Paragraph.bulletItem('First item'));
doc.addParagraph(Paragraph.bulletItem('Second item'));

// Numbered list
doc.addParagraph(Paragraph.numberedItem('Step one'));
doc.addParagraph(Paragraph.numberedItem('Step two'));

// Nested list
doc.addParagraph(Paragraph.bulletItem('Parent'));
doc.addParagraph(Paragraph.bulletItem('Child', indentLevel: 1));
```

### Tables

```dart
doc.addTable(Table(
  borders: TableBorders.all(),
  rows: [
    TableRow(cells: [
      TableCell.text('Name', backgroundColor: 'CCCCCC'),
      TableCell.text('Score', backgroundColor: 'CCCCCC'),
    ]),
    TableRow(cells: [
      TableCell.text('Alice'),
      TableCell.text('95', alignment: Alignment.right),
    ]),
  ],
));
```

### Semantic Styles

```dart
doc.addParagraph(Paragraph.heading('Title', level: 1));
doc.addParagraph(Paragraph.subtitle('Document subtitle'));
doc.addParagraph(Paragraph.quote('A famous quote...'));
doc.addParagraph(Paragraph.codeBlock('const x = 42;'));
doc.addParagraph(Paragraph.caption('Figure 1: Chart'));
```

### Page Breaks

```dart
doc.addParagraph(Paragraph.heading(
  'New Chapter',
  level: 1,
  pageBreakBefore: true,
));
```

### Line Breaks (Soft Return) - DOCX only

Line breaks allow multiple lines within a single paragraph (like Shift+Enter in Word).

```dart
// Using \n in text (automatic conversion)
doc.addParagraph(Paragraph.text('Line 1\nLine 2\nLine 3'));

// Using explicit line break runs (for different formatting per line)
doc.addParagraph(Paragraph(
  runs: [
    TextRun('Bold line', bold: true),
    TextRun.lineBreak(),
    TextRun('Normal line'),
    TextRun.lineBreak(),
    TextRun('Italic line', italic: true),
  ],
));
```

### Hyperlinks

```dart
// External link
doc.addParagraph(Paragraph(
  runs: [
    TextRun('Visit '),
    TextRun('our website', hyperlink: 'https://example.com'),
    TextRun(' for more info.'),
  ],
));
```

### Internal Links (Bookmarks)

```dart
// Create a bookmark
doc.addParagraph(Paragraph.heading(
  'Chapter 1: Introduction',
  level: 1,
  bookmarkName: 'chapter1',
));

// Link to the bookmark
doc.addParagraph(Paragraph(
  runs: [
    TextRun('Go to '),
    TextRun('Chapter 1', bookmarkRef: 'chapter1'),
  ],
));
```

### Table of Contents

```dart
// Enable automatic Table of Contents
final doc = Document(
  title: 'My Document',
  includeTableOfContents: true,
  tocTitle: 'Contents',
  tocMaxLevel: 3,  // Include Heading 1-3
);

doc.addParagraph(Paragraph.heading('Introduction', level: 1));
doc.addParagraph(Paragraph.heading('Getting Started', level: 2));
// TOC will be auto-generated with links to these headings
```

## Platform Support

| Platform | Support |
|----------|---------|
| Android | ✅ |
| iOS | ✅ |
| Web | ✅ |
| macOS | ✅ |
| Windows | ✅ |
| Linux | ✅ |

### Web Usage

```dart
import 'dart:html' as html;

void downloadDocument(Uint8List bytes, String filename) {
  final blob = html.Blob([bytes]);
  final url = html.Url.createObjectUrlFromBlob(blob);
  html.AnchorElement(href: url)
    ..setAttribute('download', filename)
    ..click();
  html.Url.revokeObjectUrl(url);
}
```

### Mobile Usage

```dart
import 'package:path_provider/path_provider.dart';
import 'package:share_plus/share_plus.dart';

Future<void> shareDocument(Uint8List bytes) async {
  final dir = await getTemporaryDirectory();
  final file = File('${dir.path}/document.docx');
  await file.writeAsBytes(bytes);
  await Share.shareXFiles([XFile(file.path)]);
}
```

## API Reference

### Main Classes

| Class | Description |
|-------|-------------|
| `Document` / `DocxDocument` | Document container with metadata |
| `Paragraph` / `DocxParagraph` | Paragraph with text runs and styling |
| `TextRun` / `DocxRun` | Text segment with formatting |
| `Table` / `DocxTable` | Table with rows and borders |
| `TableRow` / `DocxTableRow` | Table row with cells |
| `TableCell` / `DocxTableCell` | Table cell with content |
| `DocxGenerator` | Generates DOCX bytes |
| `PdfGenerator` | Generates PDF bytes |

### Paragraph Styles

| Style | Method |
|-------|--------|
| Normal text | `Paragraph.text('...')` |
| Heading 1-4 | `Paragraph.heading('...', level: 1)` |
| Subtitle | `Paragraph.subtitle('...')` |
| Caption | `Paragraph.caption('...')` |
| Quote | `Paragraph.quote('...')` |
| Code block | `Paragraph.codeBlock('...')` |
| Footnote | `Paragraph.footnote('...')` |
| Bullet list | `Paragraph.bulletItem('...')` |
| Dash list | `Paragraph.dashItem('...')` |
| Numbered list | `Paragraph.numberedItem('...')` |
| Alpha list | `Paragraph.alphaItem('...')` |
| Roman list | `Paragraph.romanItem('...')` |

### Text Formatting

| Property | Type | Description |
|----------|------|-------------|
| `bold` | `bool` | Bold text |
| `italic` | `bool` | Italic text |
| `underline` | `bool` | Underlined text |
| `strikethrough` | `bool` | Strikethrough text |
| `color` | `String` | Hex color (e.g., `'FF0000'`) |
| `backgroundColor` | `String` | Highlight color |
| `hyperlink` | `String?` | External URL link |
| `bookmarkRef` | `String?` | Internal bookmark reference |
| `isLineBreak` | `bool` | Line break (use `TextRun.lineBreak()`) |

## Compatibility

Generated documents are compatible with:

- Microsoft Word 2007+
- Google Docs
- LibreOffice Writer
- Apple Pages
- WPS Office
- Any OOXML-compatible application

## Requirements

- Dart SDK: `>=3.0.0 <4.0.0`
- Flutter: `>=3.0.0` (if using with Flutter)

## Contributing

Contributions are welcome! Please feel free to submit issues and pull requests on [GitHub](https://github.com/erykkruk/docs_gee).

## License

MIT License - see [LICENSE](LICENSE) for details.

[![Codigee - Best Flutter Experts](doc/logo.jpeg)](https://umami.team.codigee.com/q/FZ9PQYyYN)

---

<p align="center">
  Made with ❤️ by <a href="https://umami.team.codigee.com/q/FZ9PQYyYN">Codigee</a>
</p>

[![Codigee - Best Flutter Experts](logo.jpeg)](https://codigee.com)

# docs_gee API Reference

Complete API documentation for the docs_gee library.

## Table of Contents

- [Generators](#generators)
  - [DocxGenerator](#docxgenerator)
  - [PdfGenerator](#pdfgenerator)
- [Document Model](#document-model)
  - [Document](#document)
  - [Paragraph](#paragraph)
  - [TextRun](#textrun)
- [Tables](#tables)
  - [Table](#table)
  - [TableRow](#tablerow)
  - [TableCell](#tablecell)
  - [TableBorders](#tableborders)
- [Enums](#enums)
  - [Alignment](#alignment)
  - [ParagraphStyle](#paragraphstyle)
  - [BorderStyle](#borderstyle)
- [Type Aliases](#type-aliases)
- [Usage Examples](#usage-examples)

---

## Generators

### DocxGenerator

Generates Microsoft Word DOCX files.

```dart
class DocxGenerator implements DocumentGenerator
```

#### Constructor

```dart
DocxGenerator({
  String fontName = 'Times New Roman',
  int fontSize = 24,
})
```

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `fontName` | `String` | `'Times New Roman'` | Default font for the document |
| `fontSize` | `int` | `24` | Font size in half-points (24 = 12pt) |

#### Font Size Reference

| Half-points | Points | Typical Use |
|-------------|--------|-------------|
| 20 | 10pt | Footnotes |
| 22 | 11pt | Body text (compact) |
| 24 | 12pt | Body text (standard) |
| 28 | 14pt | Subheadings |
| 32 | 16pt | Section headings |
| 48 | 24pt | Titles |

#### Methods

##### generate()

```dart
Uint8List generate(Document document)
```

Returns the document as DOCX bytes.

---

### PdfGenerator

Generates PDF documents.

```dart
class PdfGenerator implements DocumentGenerator
```

#### Constructor

```dart
PdfGenerator({
  String fontName = 'Helvetica',
  double fontSize = 12.0,
})
```

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `fontName` | `String` | `'Helvetica'` | Base font (Helvetica, Times-Roman, Courier) |
| `fontSize` | `double` | `12.0` | Font size in points |

#### Character Support

PDF uses WinAnsi encoding (Windows-1252), which supports:

| Category | Characters | Support |
|----------|------------|---------|
| ASCII | A-Z, a-z, 0-9, symbols | Full |
| Typography | bullet (•), en-dash (–), em-dash (—), smart quotes, ellipsis (…), euro (€), ™, ©, ® | Full |
| German | Ä, Ö, Ü, ä, ö, ü, ß | Full |
| French | À, Â, Ç, È, É, Ê, Ë, Î, Ï, Ô, Ù, Û, etc. | Full |
| Polish | Ó, ó | Full |
| Polish | ą, ę, ć, ź, ż, ń, ł, ś | Fallback to ASCII* |

*Due to WinAnsi encoding limitations, Polish characters not in the standard set are converted to their base ASCII equivalents (e.g., ą → a, ł → l).

#### Methods

##### generate()

```dart
Uint8List generate(Document document)
```

Returns the document as PDF bytes.

---

## Document Model

### Document

Container for document content and metadata. Alias for `DocxDocument`.

```dart
Document({
  String? title,
  String? author,
})
```

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `title` | `String?` | `null` | Document title (metadata) |
| `author` | `String?` | `null` | Document author (metadata) |

#### Properties

| Property | Type | Description |
|----------|------|-------------|
| `content` | `List<Object>` | All content items (paragraphs and tables) |
| `paragraphs` | `List<DocxParagraph>` | Only paragraphs (for backward compatibility) |
| `title` | `String?` | Document title |
| `author` | `String?` | Document author |

#### Methods

```dart
void addParagraph(Paragraph paragraph)
void addParagraphs(List<Paragraph> paragraphs)
void addTable(Table table)
```

---

### Paragraph

Represents a paragraph with text content, style, and formatting. Alias for `DocxParagraph`.

#### Constructor

```dart
Paragraph({
  required List<TextRun> runs,
  ParagraphStyle style = ParagraphStyle.normal,
  Alignment alignment = Alignment.left,
  bool pageBreakBefore = false,
  int indentLevel = 0,
})
```

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `runs` | `List<TextRun>` | required | Text segments with formatting |
| `style` | `ParagraphStyle` | `normal` | Paragraph style |
| `alignment` | `Alignment` | `left` | Text alignment |
| `pageBreakBefore` | `bool` | `false` | Insert page break before |
| `indentLevel` | `int` | `0` | Nesting level for lists (0-8) |

#### Factory Constructors

```dart
// Plain text
Paragraph.text(String text, {Alignment alignment, bool pageBreakBefore})

// Headings (level 1-4)
Paragraph.heading(String text, {required int level, Alignment alignment, bool pageBreakBefore})

// Semantic styles
Paragraph.subtitle(String text, {Alignment alignment})
Paragraph.caption(String text, {Alignment alignment})
Paragraph.quote(String text, {Alignment alignment})
Paragraph.codeBlock(String text, {Alignment alignment})
Paragraph.footnote(String text, {Alignment alignment})

// Lists (with optional indentLevel for nesting)
Paragraph.bulletItem(String text, {int indentLevel, Alignment alignment})
Paragraph.dashItem(String text, {int indentLevel, Alignment alignment})
Paragraph.numberedItem(String text, {int indentLevel, Alignment alignment})
Paragraph.alphaItem(String text, {int indentLevel, Alignment alignment})
Paragraph.romanItem(String text, {int indentLevel, Alignment alignment})
```

---

### TextRun

A run of text with formatting. Alias for `DocxRun`.

```dart
TextRun(
  String text, {
  bool bold = false,
  bool italic = false,
  bool underline = false,
  bool strikethrough = false,
  String? color,
  String? backgroundColor,
})
```

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `text` | `String` | required | The text content |
| `bold` | `bool` | `false` | Bold formatting |
| `italic` | `bool` | `false` | Italic formatting |
| `underline` | `bool` | `false` | Underline formatting |
| `strikethrough` | `bool` | `false` | Strikethrough formatting |
| `color` | `String?` | `null` | Text color (hex, e.g., `'FF0000'`) |
| `backgroundColor` | `String?` | `null` | Highlight color (hex) |

#### Properties

| Property | Type | Description |
|----------|------|-------------|
| `hasFormatting` | `bool` | True if any formatting is applied |

#### Methods

```dart
TextRun copyWith({String? text, bool? bold, bool? italic, ...})
```

---

## Tables

### Table

Table container with rows and borders. Alias for `DocxTable`.

```dart
Table({
  required List<TableRow> rows,
  TableBorders borders = const TableBorders.none(),
})
```

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `rows` | `List<TableRow>` | required | Table rows |
| `borders` | `TableBorders` | `TableBorders.none()` | Border configuration |

#### Properties

| Property | Type | Description |
|----------|------|-------------|
| `columnCount` | `int` | Number of columns (from first row) |

---

### TableRow

A row containing cells. Alias for `DocxTableRow`.

```dart
TableRow({
  required List<TableCell> cells,
})
```

---

### TableCell

A cell with content. Alias for `DocxTableCell`.

```dart
TableCell({
  required List<Paragraph> paragraphs,
  String? backgroundColor,
  Alignment alignment = Alignment.left,
})
```

#### Factory Constructor

```dart
// Simple text cell
TableCell.text(
  String text, {
  String? backgroundColor,
  Alignment alignment = Alignment.left,
})
```

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `paragraphs` | `List<Paragraph>` | required | Cell content |
| `backgroundColor` | `String?` | `null` | Fill color (hex, e.g., `'E0E0E0'`) |
| `alignment` | `Alignment` | `left` | Text alignment |

---

### TableBorders

Border configuration for tables. Alias for `DocxTableBorders`.

```dart
// All borders (single line)
const TableBorders.all({
  String color = '000000',
  int size = 4,
  BorderStyle style = BorderStyle.single,
})

// No borders
const TableBorders.none()

// Outside borders only
const TableBorders.outside({
  String color = '000000',
  int size = 4,
  BorderStyle style = BorderStyle.single,
})

// Custom borders
TableBorders({
  Border? top,
  Border? left,
  Border? bottom,
  Border? right,
  Border? insideH,
  Border? insideV,
})
```

---

## Enums

### Alignment

Text alignment options. Alias for `DocxAlignment`.

| Value | Description |
|-------|-------------|
| `left` | Left-aligned (default) |
| `center` | Centered |
| `right` | Right-aligned |
| `justify` | Justified |

### ParagraphStyle

Paragraph styles. Alias for `DocxParagraphStyle`.

| Value | Description |
|-------|-------------|
| `normal` | Standard body text |
| `heading1` | Primary heading (largest) |
| `heading2` | Secondary heading |
| `heading3` | Tertiary heading |
| `heading4` | Quaternary heading (smallest) |
| `subtitle` | Document subtitle |
| `caption` | Image/figure caption |
| `quote` | Block quote |
| `codeBlock` | Code block (monospace) |
| `footnote` | Footnote text |
| `listBullet` | Bullet list item |
| `listDash` | Dash list item |
| `listNumber` | Numbered list item |
| `listNumberAlpha` | Alphabetic list item |
| `listNumberRoman` | Roman numeral list item |

### BorderStyle

Border line styles. Alias for `DocxBorderStyle`.

| Value | Description |
|-------|-------------|
| `single` | Single line |
| `double` | Double line |
| `dashed` | Dashed line |
| `dotted` | Dotted line |
| `thick` | Thick line |

---

## Type Aliases

For cleaner, format-agnostic code:

| Alias | Original Class |
|-------|----------------|
| `Document` | `DocxDocument` |
| `Paragraph` | `DocxParagraph` |
| `TextRun` | `DocxRun` |
| `Alignment` | `DocxAlignment` |
| `ParagraphStyle` | `DocxParagraphStyle` |
| `Table` | `DocxTable` |
| `TableRow` | `DocxTableRow` |
| `TableCell` | `DocxTableCell` |
| `TableBorders` | `DocxTableBorders` |
| `Border` | `DocxBorder` |
| `BorderStyle` | `DocxBorderStyle` |

---

## Usage Examples

### Complete Document with Tables

```dart
import 'dart:io';
import 'package:docs_gee/docs_gee.dart';

void main() {
  final doc = Document(title: 'Report', author: 'Team');

  // Title
  doc.addParagraph(Paragraph.heading('Quarterly Report', level: 1));
  doc.addParagraph(Paragraph.subtitle('Q4 2024 Summary'));

  // Content
  doc.addParagraph(Paragraph.text(
    'This report summarizes performance metrics.',
    alignment: Alignment.justify,
  ));

  // Table
  doc.addTable(Table(
    borders: const TableBorders.all(),
    rows: [
      TableRow(cells: [
        TableCell.text('Metric', backgroundColor: 'E0E0E0'),
        TableCell.text('Value', backgroundColor: 'E0E0E0'),
      ]),
      TableRow(cells: [
        TableCell.text('Revenue'),
        TableCell.text('\$1.2M', alignment: Alignment.right),
      ]),
      TableRow(cells: [
        TableCell.text('Growth'),
        TableCell.text('+15%', alignment: Alignment.right),
      ]),
    ],
  ));

  // Lists
  doc.addParagraph(Paragraph.heading('Key Points', level: 2));
  doc.addParagraph(Paragraph.bulletItem('All targets met'));
  doc.addParagraph(Paragraph.bulletItem('New markets opened'));
  doc.addParagraph(Paragraph.bulletItem('Sub-item', indentLevel: 1));

  // Generate both formats
  File('report.docx').writeAsBytesSync(DocxGenerator().generate(doc));
  File('report.pdf').writeAsBytesSync(PdfGenerator().generate(doc));
}
```

### Rich Text Formatting

```dart
doc.addParagraph(Paragraph(
  runs: [
    TextRun('Normal, '),
    TextRun('bold, ', bold: true),
    TextRun('italic, ', italic: true),
    TextRun('red text, ', color: 'FF0000'),
    TextRun('highlighted', backgroundColor: 'FFFF00'),
  ],
));
```

### Nested Lists

```dart
doc.addParagraph(Paragraph.numberedItem('First item'));
doc.addParagraph(Paragraph.alphaItem('Sub-item a', indentLevel: 1));
doc.addParagraph(Paragraph.alphaItem('Sub-item b', indentLevel: 1));
doc.addParagraph(Paragraph.romanItem('Detail i', indentLevel: 2));
doc.addParagraph(Paragraph.numberedItem('Second item'));
```

### Semantic Styles

```dart
doc.addParagraph(Paragraph.heading('Title', level: 1));
doc.addParagraph(Paragraph.subtitle('Document subtitle'));
doc.addParagraph(Paragraph.quote('A famous quote...'));
doc.addParagraph(Paragraph.codeBlock('const x = 42;'));
doc.addParagraph(Paragraph.caption('Figure 1: Chart'));
doc.addParagraph(Paragraph.footnote('1. Reference note'));
```

---

## Platform Notes

### Emoji Support

Emoji characters work in **DOCX only** (Word handles them natively):

```dart
doc.addParagraph(Paragraph.text('Hello World! 👋🌍'));
```

PDF uses standard fonts without emoji support.

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

[![Codigee - Best Flutter Experts](doc/logo.jpeg)](https://umami.team.codigee.com/q/FZ9PQYyYN)

# docs_gee API Reference

Complete API documentation for the docs_gee library.

## Table of Contents

- [DocxGenerator](#docxgenerator)
- [DocxDocument](#docxdocument)
- [DocxParagraph](#docxparagraph)
- [DocxRun](#docxrun)
- [DocxAlignment](#docxalignment)
- [DocxParagraphStyle](#docxparagraphstyle)
- [Usage Examples](#usage-examples)
- [Error Handling](#error-handling)

---

## DocxGenerator

Main class for generating DOCX files from a document model.

### Constructor

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

### Font Size Reference

| Half-points | Points | Typical Use |
|-------------|--------|-------------|
| 20 | 10pt | Footnotes |
| 22 | 11pt | Body text (compact) |
| 24 | 12pt | Body text (standard) |
| 28 | 14pt | Subheadings |
| 32 | 16pt | Section headings |
| 48 | 24pt | Titles |

### Methods

#### generate()

Generates a DOCX file from the given document.

```dart
Uint8List generate(DocxDocument document)
```

| Parameter | Type | Description |
|-----------|------|-------------|
| `document` | `DocxDocument` | The document to convert |

**Returns:** `Uint8List` - The DOCX file as bytes.

**Example:**

```dart
final generator = DocxGenerator(
  fontName: 'Arial',
  fontSize: 24, // 12pt
);

final bytes = generator.generate(document);
await File('output.docx').writeAsBytes(bytes);
```

---

## DocxDocument

Container for document content and metadata.

### Constructor

```dart
DocxDocument({
  List<DocxParagraph>? paragraphs,
  String? title,
  String? author,
})
```

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `paragraphs` | `List<DocxParagraph>?` | `[]` | Initial paragraphs |
| `title` | `String?` | `null` | Document title (metadata) |
| `author` | `String?` | `null` | Document author (metadata) |

### Properties

| Property | Type | Description |
|----------|------|-------------|
| `paragraphs` | `List<DocxParagraph>` | All paragraphs in the document |
| `title` | `String?` | Document title |
| `author` | `String?` | Document author |

### Methods

#### addParagraph()

Adds a single paragraph to the document.

```dart
void addParagraph(DocxParagraph paragraph)
```

#### addParagraphs()

Adds multiple paragraphs to the document.

```dart
void addParagraphs(List<DocxParagraph> paragraphs)
```

**Example:**

```dart
final doc = DocxDocument(
  title: 'Annual Report',
  author: 'Finance Team',
);

doc.addParagraph(DocxParagraph.heading('Introduction', level: 1));
doc.addParagraphs([
  DocxParagraph.text('First paragraph...'),
  DocxParagraph.text('Second paragraph...'),
]);
```

---

## DocxParagraph

Represents a paragraph with text content, style, and formatting.

### Constructor

```dart
const DocxParagraph({
  required List<DocxRun> runs,
  DocxParagraphStyle style = DocxParagraphStyle.normal,
  DocxAlignment alignment = DocxAlignment.left,
  bool pageBreakBefore = false,
})
```

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `runs` | `List<DocxRun>` | required | Text segments with formatting |
| `style` | `DocxParagraphStyle` | `normal` | Paragraph style |
| `alignment` | `DocxAlignment` | `left` | Text alignment |
| `pageBreakBefore` | `bool` | `false` | Insert page break before this paragraph |

### Factory Constructors

#### DocxParagraph.text()

Creates a simple paragraph with plain text.

```dart
factory DocxParagraph.text(
  String text, {
  DocxParagraphStyle style = DocxParagraphStyle.normal,
  DocxAlignment alignment = DocxAlignment.left,
  bool pageBreakBefore = false,
})
```

#### DocxParagraph.heading()

Creates a heading paragraph.

```dart
factory DocxParagraph.heading(
  String text, {
  required int level,  // 1, 2, or 3
  DocxAlignment alignment = DocxAlignment.left,
  bool pageBreakBefore = false,
})
```

| Level | Style Applied |
|-------|---------------|
| 1 | `heading1` (largest) |
| 2 | `heading2` (medium) |
| 3 | `heading3` (smallest) |

#### DocxParagraph.bulletItem()

Creates a bullet list item.

```dart
factory DocxParagraph.bulletItem(
  String text, {
  DocxAlignment alignment = DocxAlignment.left,
})
```

#### DocxParagraph.numberedItem()

Creates a numbered list item.

```dart
factory DocxParagraph.numberedItem(
  String text, {
  DocxAlignment alignment = DocxAlignment.left,
})
```

### Properties

| Property | Type | Description |
|----------|------|-------------|
| `runs` | `List<DocxRun>` | Text segments in this paragraph |
| `style` | `DocxParagraphStyle` | The paragraph style |
| `alignment` | `DocxAlignment` | Text alignment |
| `pageBreakBefore` | `bool` | Whether to break page before |
| `plainText` | `String` | Combined text from all runs (getter) |

**Example:**

```dart
// Simple text
doc.addParagraph(DocxParagraph.text('Hello World'));

// Heading with page break
doc.addParagraph(DocxParagraph.heading(
  'Chapter 2',
  level: 1,
  pageBreakBefore: true,
));

// Mixed formatting using runs
doc.addParagraph(DocxParagraph(
  runs: [
    const DocxRun('Regular '),
    const DocxRun('bold ', bold: true),
    const DocxRun('italic', italic: true),
  ],
  alignment: DocxAlignment.center,
));
```

---

## DocxRun

A run of text with consistent formatting. In DOCX terminology, a "run" is a contiguous piece of text sharing the same formatting properties.

### Constructor

```dart
const DocxRun(
  String text, {
  bool bold = false,
  bool italic = false,
  bool underline = false,
  bool strikethrough = false,
})
```

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `text` | `String` | required | The text content |
| `bold` | `bool` | `false` | Bold formatting |
| `italic` | `bool` | `false` | Italic formatting |
| `underline` | `bool` | `false` | Underline formatting |
| `strikethrough` | `bool` | `false` | Strikethrough formatting |

### Properties

| Property | Type | Description |
|----------|------|-------------|
| `text` | `String` | The text content |
| `bold` | `bool` | Whether text is bold |
| `italic` | `bool` | Whether text is italic |
| `underline` | `bool` | Whether text is underlined |
| `strikethrough` | `bool` | Whether text has strikethrough |
| `hasFormatting` | `bool` | True if any formatting is applied (getter) |

### Methods

#### copyWith()

Creates a copy with modified properties.

```dart
DocxRun copyWith({
  String? text,
  bool? bold,
  bool? italic,
  bool? underline,
  bool? strikethrough,
})
```

**Example:**

```dart
const original = DocxRun('Hello', bold: true);
final modified = original.copyWith(italic: true);
// Result: bold AND italic
```

**Combined Formatting Example:**

```dart
doc.addParagraph(DocxParagraph(
  runs: [
    const DocxRun('Normal, '),
    const DocxRun('bold, ', bold: true),
    const DocxRun('italic, ', italic: true),
    const DocxRun('bold+italic, ', bold: true, italic: true),
    const DocxRun('underlined, ', underline: true),
    const DocxRun('struck through.', strikethrough: true),
  ],
));
```

---

## DocxAlignment

Text alignment options for paragraphs.

```dart
enum DocxAlignment {
  left('left'),
  center('center'),
  right('right'),
  justify('both');
}
```

| Value | Word XML Value | Description |
|-------|----------------|-------------|
| `left` | `left` | Left-aligned (default) |
| `center` | `center` | Centered |
| `right` | `right` | Right-aligned |
| `justify` | `both` | Justified (both edges) |

**Example:**

```dart
doc.addParagraph(DocxParagraph.text(
  'This text is centered.',
  alignment: DocxAlignment.center,
));
```

---

## DocxParagraphStyle

Predefined paragraph styles.

```dart
enum DocxParagraphStyle {
  normal('Normal'),
  heading1('Heading1'),
  heading2('Heading2'),
  heading3('Heading3'),
  listBullet('ListBullet'),
  listNumber('ListNumber');
}
```

| Value | Style ID | Description |
|-------|----------|-------------|
| `normal` | `Normal` | Standard body text |
| `heading1` | `Heading1` | Primary heading (largest) |
| `heading2` | `Heading2` | Secondary heading |
| `heading3` | `Heading3` | Tertiary heading (smallest) |
| `listBullet` | `ListBullet` | Bullet list item |
| `listNumber` | `ListNumber` | Numbered list item |

---

## Usage Examples

### Complete Document

```dart
import 'dart:io';
import 'package:docs_gee/docs_gee.dart';

void main() {
  final doc = DocxDocument(
    title: 'Project Report',
    author: 'Development Team',
  );

  // Title page
  doc.addParagraph(DocxParagraph.heading(
    'Project Report 2024',
    level: 1,
    alignment: DocxAlignment.center,
  ));

  doc.addParagraph(DocxParagraph.text(
    'Prepared by: Development Team',
    alignment: DocxAlignment.center,
  ));

  // New section with page break
  doc.addParagraph(DocxParagraph.heading(
    'Executive Summary',
    level: 1,
    pageBreakBefore: true,
  ));

  doc.addParagraph(DocxParagraph.text(
    'This report summarizes the key achievements and challenges '
    'encountered during the project lifecycle.',
    alignment: DocxAlignment.justify,
  ));

  // Key points as bullet list
  doc.addParagraph(DocxParagraph.heading('Key Points', level: 2));
  doc.addParagraph(DocxParagraph.bulletItem('Completed all milestones'));
  doc.addParagraph(DocxParagraph.bulletItem('Under budget by 10%'));
  doc.addParagraph(DocxParagraph.bulletItem('Delivered ahead of schedule'));

  // Action items as numbered list
  doc.addParagraph(DocxParagraph.heading('Next Steps', level: 2));
  doc.addParagraph(DocxParagraph.numberedItem('Review final deliverables'));
  doc.addParagraph(DocxParagraph.numberedItem('Schedule stakeholder meeting'));
  doc.addParagraph(DocxParagraph.numberedItem('Prepare presentation'));

  // Generate and save
  final generator = DocxGenerator(
    fontName: 'Calibri',
    fontSize: 22, // 11pt
  );

  final bytes = generator.generate(doc);
  File('report.docx').writeAsBytesSync(bytes);
  print('Generated report.docx (${bytes.length} bytes)');
}
```

### Flutter/Web Usage

```dart
import 'package:docs_gee/docs_gee.dart';

class DocumentService {
  Uint8List generateInvoice({
    required String clientName,
    required List<LineItem> items,
  }) {
    final doc = DocxDocument(title: 'Invoice');

    doc.addParagraph(DocxParagraph.heading('INVOICE', level: 1));
    doc.addParagraph(DocxParagraph.text('Client: $clientName'));
    doc.addParagraph(DocxParagraph.text(''));

    for (final item in items) {
      doc.addParagraph(DocxParagraph(
        runs: [
          DocxRun(item.name, bold: true),
          DocxRun(' - \$${item.price.toStringAsFixed(2)}'),
        ],
      ));
    }

    final generator = DocxGenerator();
    return generator.generate(doc);
  }
}
```

---

## Error Handling

The library uses standard Dart exceptions:

| Scenario | Exception |
|----------|-----------|
| Empty document | No exception (generates valid empty document) |
| Invalid UTF-8 text | `FormatException` from dart:convert |
| Memory issues | Standard Dart memory errors |

**Best Practice:**

```dart
try {
  final bytes = generator.generate(doc);
  await file.writeAsBytes(bytes);
} catch (e) {
  print('Failed to generate document: $e');
}
```

---

## Performance Tips

1. **Reuse generator** - Create one `DocxGenerator` instance and reuse it for multiple documents with the same font settings.

2. **Batch paragraphs** - Use `addParagraphs()` instead of multiple `addParagraph()` calls when adding many paragraphs at once.

3. **Minimize runs** - Combine adjacent text with the same formatting into a single `DocxRun` instead of multiple runs.

```dart
// Less efficient
runs: [
  DocxRun('Hello '),
  DocxRun('World'),
]

// More efficient
runs: [
  DocxRun('Hello World'),
]
```

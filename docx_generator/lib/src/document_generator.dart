import 'dart:typed_data';

import 'models/models.dart';

/// Abstract interface for document generators.
///
/// Both [DocxGenerator] and [PdfGenerator] implement this interface,
/// allowing them to be used interchangeably.
///
/// This interface only contains platform-independent methods.
/// For file I/O operations (not available on Web), use the concrete
/// generator classes directly which provide [generateToFile] method.
///
/// Example:
/// ```dart
/// final doc = Document();
/// doc.addParagraph(Paragraph.text('Hello World'));
///
/// // Use any generator (works on all platforms including Web)
/// DocumentGenerator generator = DocxGenerator();
/// final bytes = generator.generate(doc);
///
/// // For file saving (not available on Web):
/// // await DocxGenerator().generateToFile(doc, filePath: 'output.docx');
/// ```
abstract class DocumentGenerator {
  /// Generates the document and returns it as bytes.
  ///
  /// Works on all platforms including Web.
  Uint8List generate(DocxDocument document);
}

// ============================================================
// Type aliases for generic naming
// ============================================================

/// Alias for [DocxDocument] - use for format-agnostic code.
typedef Document = DocxDocument;

/// Alias for [DocxParagraph] - use for format-agnostic code.
typedef Paragraph = DocxParagraph;

/// Alias for [DocxRun] - use for format-agnostic code.
typedef TextRun = DocxRun;

/// Alias for [DocxAlignment] - use for format-agnostic code.
typedef Alignment = DocxAlignment;

/// Alias for [DocxParagraphStyle] - use for format-agnostic code.
typedef ParagraphStyle = DocxParagraphStyle;

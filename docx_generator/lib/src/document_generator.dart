import 'dart:typed_data';

import 'models/models.dart';

/// Abstract interface for document generators.
///
/// Both [DocxGenerator] and [PdfGenerator] implement this interface,
/// allowing them to be used interchangeably.
///
/// Example:
/// ```dart
/// final doc = Document();
/// doc.addParagraph(Paragraph.text('Hello World'));
///
/// // Use any generator
/// DocumentGenerator generator = DocxGenerator();
/// // or: DocumentGenerator generator = PdfGenerator();
///
/// final bytes = generator.generate(doc);
///
/// // Or save directly to file
/// await generator.generateToFile(doc, filePath: 'my_document.docx');
/// ```
abstract class DocumentGenerator {
  /// Generates the document and returns it as bytes.
  Uint8List generate(DocxDocument document);

  /// Generates the document and saves it to a file.
  ///
  /// [filePath] - optional path where to save the file.
  /// If not provided, uses a default name (e.g., 'document.docx' or 'document.pdf').
  ///
  /// Returns the actual file path where the document was saved.
  Future<String> generateToFile(DocxDocument document, {String? filePath});

  /// Default file extension for this generator (e.g., '.docx', '.pdf').
  String get defaultExtension;
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

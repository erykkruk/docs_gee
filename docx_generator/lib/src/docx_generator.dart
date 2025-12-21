import 'dart:convert';
import 'dart:typed_data';

import 'package:archive/archive.dart';

import 'document_generator.dart';
import 'models/models.dart';
import 'xml_parts/xml_parts.dart';

/// Generator for creating DOCX files.
///
/// Implements [DocumentGenerator] interface for interchangeable use with [PdfGenerator].
///
/// Example usage:
/// ```dart
/// final doc = DocxDocument();
/// doc.addParagraph(DocxParagraph.heading('My Title', level: 1));
/// doc.addParagraph(DocxParagraph.text('Hello world!'));
///
/// final generator = DocxGenerator();
/// final bytes = generator.generate(doc);
/// await File('output.docx').writeAsBytes(bytes);
/// ```
class DocxGenerator implements DocumentGenerator {
  /// Creates a new DOCX generator.
  ///
  /// [fontName] - default font name for the document.
  /// [fontSize] - default font size in half-points (24 = 12pt).
  DocxGenerator({
    this.fontName = 'Times New Roman',
    this.fontSize = 24,
  });

  /// Default font name for the document.
  final String fontName;

  /// Default font size in half-points (24 = 12pt, 28 = 14pt).
  final int fontSize;

  /// Default file extension for DOCX files.
  static const String defaultExtension = '.docx';

  /// Generates a DOCX file from the given document.
  ///
  /// Returns the DOCX file as bytes that can be written to a file.
  @override
  Uint8List generate(DocxDocument document) {
    final hasLists = _documentHasLists(document);

    // Generate document.xml and collect hyperlinks
    final documentResult = DocumentXml.generate(document);

    final archive = Archive();

    // Add [Content_Types].xml
    _addFile(
      archive,
      '[Content_Types].xml',
      ContentTypesXml.generate(hasNumbering: hasLists),
    );

    // Add _rels/.rels
    _addFile(
      archive,
      '_rels/.rels',
      RelsXml.generateMainRels(),
    );

    // Add word/_rels/document.xml.rels (with hyperlinks)
    _addFile(
      archive,
      'word/_rels/document.xml.rels',
      RelsXml.generateDocumentRels(
        hasNumbering: hasLists,
        hyperlinks: documentResult.hyperlinks,
      ),
    );

    // Add word/document.xml
    _addFile(
      archive,
      'word/document.xml',
      documentResult.xml,
    );

    // Add word/styles.xml
    _addFile(
      archive,
      'word/styles.xml',
      StylesXml.generate(fontName: fontName, fontSize: fontSize),
    );

    // Add word/numbering.xml if needed
    if (hasLists) {
      _addFile(
        archive,
        'word/numbering.xml',
        NumberingXml.generate(),
      );
    }

    // Encode as ZIP
    final zipEncoder = ZipEncoder();
    final zipBytes = zipEncoder.encode(archive);

    return Uint8List.fromList(zipBytes);
  }

  void _addFile(Archive archive, String path, String content) {
    final bytes = utf8.encode(content);
    archive.addFile(ArchiveFile(path, bytes.length, bytes));
  }

  bool _documentHasLists(DocxDocument document) {
    return document.paragraphs.any(
      (p) =>
          p.style == DocxParagraphStyle.listBullet ||
          p.style == DocxParagraphStyle.listNumber,
    );
  }
}

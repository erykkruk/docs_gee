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

  /// Relationship ID of word/header1.xml.
  static const String _headerRelId = 'rId3';

  /// Relationship ID of word/footer1.xml.
  static const String _footerRelId = 'rId4';

  /// Generates a DOCX file from the given document.
  ///
  /// Returns the DOCX file as bytes that can be written to a file.
  @override
  Uint8List generate(DocxDocument document) {
    final hasLists = _documentHasLists(document);
    final header = _partWithContent(document.header);
    final footer = _partWithContent(document.footer);
    final headerRelId = header == null ? null : _headerRelId;
    final footerRelId = footer == null ? null : _footerRelId;

    // Generate document.xml and collect hyperlink and image relationships
    final documentResult = DocumentXml.generate(
      document,
      headerRelId: headerRelId,
      footerRelId: footerRelId,
    );
    final imageFormats = document.images.map((image) => image.format).toSet();

    final archive = Archive();

    // Add [Content_Types].xml
    _addFile(
      archive,
      '[Content_Types].xml',
      ContentTypesXml.generate(
        hasNumbering: hasLists,
        imageFormats: imageFormats,
        hasHeader: header != null,
        hasFooter: footer != null,
      ),
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
        images: documentResult.images,
        headerRelId: headerRelId,
        footerRelId: footerRelId,
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

    // Add word/header1.xml and word/footer1.xml if defined
    if (header != null) {
      _addFile(
        archive,
        'word/header1.xml',
        HeaderFooterXml.generateHeader(header),
      );
    }
    if (footer != null) {
      _addFile(
        archive,
        'word/footer1.xml',
        HeaderFooterXml.generateFooter(footer),
      );
    }

    // Add the media parts backing the image relationships. Iterating the
    // relationship map rather than document.images keeps part names and
    // r:embed ids in lockstep with what document.xml actually referenced.
    final images = document.images;
    var imageIndex = 0;
    for (final partName in documentResult.images.values) {
      final image = images[imageIndex];
      imageIndex++;
      _addBinaryFile(archive, 'word/$partName', image.bytes);
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

  void _addBinaryFile(Archive archive, String path, Uint8List bytes) {
    archive.addFile(ArchiveFile(path, bytes.length, bytes));
  }

  /// Returns the header or footer only when it would render something, so an
  /// empty one never adds a part or a relationship to the archive.
  DocxHeaderFooter? _partWithContent(DocxHeaderFooter? part) {
    if (part == null || part.isEmpty) return null;
    return part;
  }

  bool _documentHasLists(DocxDocument document) {
    return document.paragraphs.any((p) => p.style.isList);
  }
}

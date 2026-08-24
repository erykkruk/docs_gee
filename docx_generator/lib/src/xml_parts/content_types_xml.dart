import '../models/models.dart';
import 'xml_utils.dart';

/// Generates [Content_Types].xml for DOCX.
class ContentTypesXml {
  ContentTypesXml._();

  /// Generates the [Content_Types].xml content.
  ///
  /// [hasNumbering] - whether to include numbering.xml (for lists).
  /// [imageFormats] - raster formats embedded in the document; each one needs
  /// a Default extension entry or Word rejects the archive.
  /// [hasHeader] / [hasFooter] - whether header1.xml / footer1.xml are part
  /// of the archive.
  static String generate({
    bool hasNumbering = false,
    Set<DocxImageFormat> imageFormats = const {},
    bool hasHeader = false,
    bool hasFooter = false,
  }) {
    final buffer = StringBuffer();
    buffer.writeln(XmlUtils.xmlDeclaration);
    buffer.writeln('<Types xmlns="${XmlUtils.contentTypesNamespace}">');

    // Default extensions
    buffer.writeln('  <Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>');
    buffer.writeln('  <Default Extension="xml" '
        'ContentType="application/xml"/>');

    // One Default per embedded raster format. Sorted so the output is
    // byte-stable across runs.
    final sortedFormats = imageFormats.toList()
      ..sort((a, b) => a.extension.compareTo(b.extension));
    for (final format in sortedFormats) {
      buffer.writeln('  <Default Extension="${format.extension}" '
          'ContentType="${format.mimeType}"/>');
    }

    // Override parts
    buffer.writeln('  <Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>');
    buffer.writeln('  <Override PartName="/word/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>');

    if (hasNumbering) {
      buffer.writeln('  <Override PartName="/word/numbering.xml" '
          'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>');
    }

    if (hasHeader) {
      buffer.writeln('  <Override PartName="/word/header1.xml" '
          'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>');
    }

    if (hasFooter) {
      buffer.writeln('  <Override PartName="/word/footer1.xml" '
          'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>');
    }

    buffer.writeln('</Types>');
    return buffer.toString();
  }
}

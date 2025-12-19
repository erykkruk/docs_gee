import 'xml_utils.dart';

/// Generates [Content_Types].xml for DOCX.
class ContentTypesXml {
  ContentTypesXml._();

  /// Generates the [Content_Types].xml content.
  ///
  /// [hasNumbering] - whether to include numbering.xml (for lists).
  static String generate({bool hasNumbering = false}) {
    final buffer = StringBuffer();
    buffer.writeln(XmlUtils.xmlDeclaration);
    buffer.writeln('<Types xmlns="${XmlUtils.contentTypesNamespace}">');

    // Default extensions
    buffer.writeln('  <Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>');
    buffer.writeln('  <Default Extension="xml" '
        'ContentType="application/xml"/>');

    // Override parts
    buffer.writeln('  <Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>');
    buffer.writeln('  <Override PartName="/word/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>');

    if (hasNumbering) {
      buffer.writeln('  <Override PartName="/word/numbering.xml" '
          'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>');
    }

    buffer.writeln('</Types>');
    return buffer.toString();
  }
}

import 'xml_utils.dart';

/// Generates .rels files for DOCX.
class RelsXml {
  RelsXml._();

  /// Generates the main _rels/.rels content.
  static String generateMainRels() {
    final buffer = StringBuffer();
    buffer.writeln(XmlUtils.xmlDeclaration);
    buffer.writeln('<Relationships xmlns="${XmlUtils.relsNamespace}">');
    buffer.writeln('  <Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>');
    buffer.writeln('</Relationships>');
    return buffer.toString();
  }

  /// Generates word/_rels/document.xml.rels content.
  ///
  /// [hasNumbering] - whether to include numbering.xml relationship.
  static String generateDocumentRels({bool hasNumbering = false}) {
    final buffer = StringBuffer();
    buffer.writeln(XmlUtils.xmlDeclaration);
    buffer.writeln('<Relationships xmlns="${XmlUtils.relsNamespace}">');
    buffer.writeln('  <Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>');

    if (hasNumbering) {
      buffer.writeln('  <Relationship Id="rId2" '
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" '
          'Target="numbering.xml"/>');
    }

    buffer.writeln('</Relationships>');
    return buffer.toString();
  }
}

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
  /// [hyperlinks] - map of relationship IDs to URLs for external hyperlinks.
  /// [images] - map of relationship IDs to media part names, relative to
  /// `word/` (for example `media/image1.png`).
  /// [headerRelId] / [footerRelId] - relationship IDs for header1.xml and
  /// footer1.xml, when the document defines them.
  static String generateDocumentRels({
    bool hasNumbering = false,
    Map<String, String> hyperlinks = const {},
    Map<String, String> images = const {},
    String? headerRelId,
    String? footerRelId,
  }) {
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

    if (headerRelId != null) {
      buffer.writeln('  <Relationship Id="$headerRelId" '
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" '
          'Target="header1.xml"/>');
    }

    if (footerRelId != null) {
      buffer.writeln('  <Relationship Id="$footerRelId" '
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" '
          'Target="footer1.xml"/>');
    }

    // Add image relationships
    for (final entry in images.entries) {
      buffer.writeln('  <Relationship Id="${entry.key}" '
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
          'Target="${XmlUtils.escapeXml(entry.value)}"/>');
    }

    // Add hyperlink relationships
    for (final entry in hyperlinks.entries) {
      buffer.writeln('  <Relationship Id="${entry.key}" '
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
          'Target="${XmlUtils.escapeXml(entry.value)}" '
          'TargetMode="External"/>');
    }

    buffer.writeln('</Relationships>');
    return buffer.toString();
  }
}

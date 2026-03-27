import 'package:xml/xml.dart';

import '../docx_reader_exception.dart';

/// Parses `word/_rels/document.xml.rels` to extract hyperlink relationships.
///
/// Returns a map of relationship IDs to target URLs for external hyperlinks.
class RelationshipResolver {
  const RelationshipResolver._();

  static const _hyperlinkType =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

  /// Parses the relationships XML and returns `{rId: url}` for hyperlinks.
  static Map<String, String> resolve(String xmlContent) {
    final XmlDocument document;
    try {
      document = XmlDocument.parse(xmlContent);
    } on XmlException catch (e) {
      throw InvalidDocxXmlException(
        'Failed to parse document.xml.rels: ${e.message}',
      );
    }

    final relationships = <String, String>{};
    final elements = document.findAllElements('Relationship');

    for (final element in elements) {
      final type = element.getAttribute('Type');
      if (type != _hyperlinkType) continue;

      final id = element.getAttribute('Id');
      final target = element.getAttribute('Target');
      if (id != null && target != null) {
        relationships[id] = target;
      }
    }

    return relationships;
  }
}

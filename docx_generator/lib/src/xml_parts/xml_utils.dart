/// Utility functions for XML generation.
class XmlUtils {
  XmlUtils._();

  /// Escapes special XML characters in text.
  static String escapeXml(String text) {
    return text
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&apos;');
  }

  /// XML declaration header.
  static const String xmlDeclaration =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

  /// Main WordprocessingML namespace.
  static const String wNamespace =
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

  /// Relationships namespace.
  static const String relsNamespace =
      'http://schemas.openxmlformats.org/package/2006/relationships';

  /// Content Types namespace.
  static const String contentTypesNamespace =
      'http://schemas.openxmlformats.org/package/2006/content-types';
}

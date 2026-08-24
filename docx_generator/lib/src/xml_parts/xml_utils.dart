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

  /// Relationship reference namespace (for r: prefix in documents).
  static const String rNamespace =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  /// Content Types namespace.
  static const String contentTypesNamespace =
      'http://schemas.openxmlformats.org/package/2006/content-types';

  /// WordprocessingML drawing namespace (the `wp:` prefix), which wraps an
  /// embedded picture in a way Word can lay out inline with text.
  static const String wpNamespace =
      'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';

  /// DrawingML main namespace (the `a:` prefix).
  static const String aNamespace =
      'http://schemas.openxmlformats.org/drawingml/2006/main';

  /// DrawingML picture namespace (the `pic:` prefix).
  static const String picNamespace =
      'http://schemas.openxmlformats.org/drawingml/2006/picture';
}

/// Utility functions for XML generation.
class XmlUtils {
  XmlUtils._();

  /// Characters XML 1.0 forbids anywhere in a document: C0 controls other
  /// than tab, line feed and carriage return, plus U+FFFE and U+FFFF. Even
  /// escaped as character references they make the part malformed, and Word
  /// refuses to open the file.
  static final RegExp _invalidXmlChars =
      RegExp('[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]');

  /// Escapes special XML characters in text.
  ///
  /// Characters that XML 1.0 does not allow (such as the vertical tab or
  /// NUL that text pasted from other applications often carries) are
  /// dropped, so the generated part always stays well-formed.
  static String escapeXml(String text) {
    return text
        .replaceAll(_invalidXmlChars, '')
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

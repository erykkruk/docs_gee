import 'xml_utils.dart';

/// Generates docProps/core.xml, the package part holding document metadata
/// such as the title and author that Word shows under File > Info.
class CorePropertiesXml {
  CorePropertiesXml._();

  /// Core properties namespace (the `cp:` prefix).
  static const String cpNamespace =
      'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';

  /// Dublin Core elements namespace (the `dc:` prefix).
  static const String dcNamespace = 'http://purl.org/dc/elements/1.1/';

  /// Path of the part inside the archive.
  static const String partName = 'docProps/core.xml';

  /// Whether a document with this metadata needs a core properties part.
  ///
  /// Empty strings count as unset, matching how `DocxReader` reads them, so
  /// a document without metadata keeps its archive unchanged.
  static bool hasContent({String? title, String? author}) {
    return (title?.isNotEmpty ?? false) || (author?.isNotEmpty ?? false);
  }

  /// Generates the docProps/core.xml content.
  ///
  /// [title] becomes `dc:title` and [author] becomes `dc:creator`; unset or
  /// empty values are left out.
  static String generate({String? title, String? author}) {
    final buffer = StringBuffer();
    buffer.writeln(XmlUtils.xmlDeclaration);
    buffer.writeln('<cp:coreProperties xmlns:cp="$cpNamespace" '
        'xmlns:dc="$dcNamespace">');
    if (title != null && title.isNotEmpty) {
      buffer.writeln('  <dc:title>${XmlUtils.escapeXml(title)}</dc:title>');
    }
    if (author != null && author.isNotEmpty) {
      buffer
          .writeln('  <dc:creator>${XmlUtils.escapeXml(author)}</dc:creator>');
    }
    buffer.writeln('</cp:coreProperties>');
    return buffer.toString();
  }
}

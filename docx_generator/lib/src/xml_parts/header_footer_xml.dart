import '../models/models.dart';
import 'document_xml.dart';
import 'xml_utils.dart';

/// Generates word/header1.xml and word/footer1.xml for DOCX.
///
/// Both parts share the paragraph markup of the main document, so the
/// paragraph writer is reused from [DocumentXml]. External hyperlinks are
/// rendered as plain text here: a header or footer resolves `r:id` against
/// its own relationship file, which this library does not emit, and a
/// dangling id makes Word reject the whole document.
class HeaderFooterXml {
  HeaderFooterXml._();

  /// Generates the header part content.
  static String generateHeader(DocxHeaderFooter header) {
    return _generate(header, rootElement: 'w:hdr');
  }

  /// Generates the footer part content.
  static String generateFooter(DocxHeaderFooter footer) {
    return _generate(footer, rootElement: 'w:ftr');
  }

  static String _generate(
    DocxHeaderFooter content, {
    required String rootElement,
  }) {
    final buffer = StringBuffer();
    // Discarded: header and footer parts carry no hyperlink relationships.
    final hyperlinks = <String, String>{};

    buffer.writeln(XmlUtils.xmlDeclaration);
    buffer.writeln('<$rootElement xmlns:w="${XmlUtils.wNamespace}" '
        'xmlns:r="${XmlUtils.rNamespace}">');

    if (content.paragraphs.isEmpty) {
      // A header part with no paragraph at all is invalid.
      buffer.writeln('  <w:p/>');
    } else {
      for (final paragraph in content.paragraphs) {
        DocumentXml.writeParagraph(
          buffer,
          paragraph,
          hyperlinks,
          allowHyperlinks: false,
        );
      }
    }

    buffer.writeln('</$rootElement>');
    return buffer.toString();
  }
}

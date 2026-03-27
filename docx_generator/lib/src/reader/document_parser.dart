import 'package:xml/xml.dart';

import '../docx_reader_exception.dart';
import '../models/models.dart';
import 'run_parser.dart';
import 'style_resolver.dart';
import 'table_parser.dart';

/// Parses `word/document.xml` into a list of content items
/// ([DocxParagraph] and [DocxTable]) in their original order.
class DocumentParser {
  const DocumentParser._();

  /// Parses the document.xml content.
  ///
  /// Returns a list of [DocxParagraph] and [DocxTable] objects in document order.
  static List<Object> parse(
    String documentXml,
    Map<String, String> relationships, {
    String? numberingXml,
  }) {
    final XmlDocument document;
    try {
      document = XmlDocument.parse(documentXml);
    } on XmlException catch (e) {
      throw InvalidDocxXmlException(
        'Failed to parse document.xml: ${e.message}',
      );
    }

    final documentElement = document.rootElement;
    XmlElement? body;
    for (final child in documentElement.children) {
      if (child is XmlElement && child.name.local == 'body') {
        body = child;
        break;
      }
    }
    if (body == null) {
      throw const InvalidDocxXmlException('Missing <w:body> in document.xml');
    }

    final content = <Object>[];
    var nextPageBreak = false;

    for (final child in body.children) {
      if (child is! XmlElement) continue;

      switch (child.name.local) {
        case 'p':
          final result = _parseParagraph(
            child,
            relationships,
            pageBreakBefore: nextPageBreak,
            numberingXml: numberingXml,
          );
          nextPageBreak = false;

          if (result == null) {
            // Page-break-only paragraph → apply to next paragraph
            nextPageBreak = true;
          } else {
            content.add(result);
          }

        case 'tbl':
          final table = TableParser.parse(child, relationships);
          content.add(table);
          nextPageBreak = false;

        case 'sectPr':
          // Section properties — skip
          break;
      }
    }

    return content;
  }

  /// Parses a `<w:p>` element into a [DocxParagraph].
  ///
  /// Returns `null` if the paragraph is a page-break-only paragraph
  /// (the page break should be applied to the next paragraph).
  static DocxParagraph? _parseParagraph(
    XmlElement pElement,
    Map<String, String> relationships, {
    bool pageBreakBefore = false,
    String? numberingXml,
  }) {
    final pPr = _findChild(pElement, 'pPr');

    // Check if this is a page-break-only paragraph
    if (_isPageBreakOnlyParagraph(pElement)) {
      return null;
    }

    // Style
    var style = DocxParagraphStyle.normal;
    final pStyle = _findChild(pPr, 'pStyle');
    final styleId = _getAttr(pStyle, 'val');
    style = StyleResolver.resolveStyleId(styleId);

    // Numbering (lists)
    var indentLevel = 0;
    final numPr = _findChild(pPr, 'numPr');
    if (numPr != null) {
      final ilvl = _findChild(numPr, 'ilvl');
      indentLevel = int.tryParse(_getAttr(ilvl, 'val') ?? '') ?? 0;

      final numId = _findChild(numPr, 'numId');
      final numIdVal = int.tryParse(_getAttr(numId, 'val') ?? '');
      if (numIdVal != null) {
        // numPr in paragraph overrides style-based list type
        final listStyle = StyleResolver.resolveNumId(
          numIdVal,
          numberingXml: numberingXml,
        );
        if (listStyle.isList) {
          style = listStyle;
        }
      }
    }

    // Alignment
    final jc = _findChild(pPr, 'jc');
    final alignment = StyleResolver.resolveAlignment(_getAttr(jc, 'val'));

    // Bookmark
    String? bookmarkName;
    final bookmarkStart = _findChild(pElement, 'bookmarkStart');
    if (bookmarkStart != null) {
      bookmarkName = _getAttr(bookmarkStart, 'name');
    }

    // Parse runs
    final runs = _parseRuns(pElement, relationships);

    return DocxParagraph(
      runs: runs,
      style: style,
      alignment: alignment,
      pageBreakBefore: pageBreakBefore,
      indentLevel: indentLevel,
      bookmarkName: bookmarkName,
    );
  }

  /// Checks if a paragraph is a page-break-only paragraph.
  ///
  /// docs_gee generates page breaks as separate `<w:p>` with only
  /// `<w:r><w:br w:type="page"/></w:r>` inside.
  static bool _isPageBreakOnlyParagraph(XmlElement pElement) {
    final runs = <XmlElement>[];
    for (final child in pElement.children) {
      if (child is XmlElement && child.name.local == 'r') {
        runs.add(child);
      }
    }

    if (runs.isEmpty) return false;

    // Check if all runs contain only page breaks
    for (final run in runs) {
      for (final child in run.children) {
        if (child is! XmlElement) continue;
        if (child.name.local == 'br') {
          final brType = _getAttr(child, 'type');
          if (brType != 'page') return false;
        } else if (child.name.local == 't') {
          // Has text content → not a page-break-only paragraph
          return false;
        } else if (child.name.local == 'rPr') {
          // Run properties are fine, ignore them
          continue;
        }
      }
    }

    return true;
  }

  /// Parses all runs within a paragraph element, handling hyperlinks.
  static List<DocxRun> _parseRuns(
    XmlElement pElement,
    Map<String, String> relationships,
  ) {
    final runs = <DocxRun>[];

    for (final child in pElement.children) {
      if (child is! XmlElement) continue;

      if (child.name.local == 'r') {
        // Check for page break inside a mixed paragraph (skip it)
        final br = _findChild(child, 'br');
        if (br != null && _getAttr(br, 'type') == 'page') continue;

        runs.add(RunParser.parse(child));
      } else if (child.name.local == 'hyperlink') {
        _parseHyperlinkRuns(child, relationships, runs);
      }
    }

    return runs;
  }

  /// Parses runs inside a `<w:hyperlink>` element.
  static void _parseHyperlinkRuns(
    XmlElement hyperlinkElement,
    Map<String, String> relationships,
    List<DocxRun> runs,
  ) {
    // External: <w:hyperlink r:id="rId100">
    final rId = hyperlinkElement.getAttribute('r:id') ??
        hyperlinkElement.getAttribute('id',
            namespace:
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
    // Internal: <w:hyperlink w:anchor="bookmarkName">
    final anchor = _getAttr(hyperlinkElement, 'anchor');

    String? hyperlinkUrl;
    String? bookmarkRef;

    if (rId != null) {
      hyperlinkUrl = relationships[rId];
    } else if (anchor != null) {
      bookmarkRef = anchor;
    }

    for (final hChild in hyperlinkElement.children) {
      if (hChild is XmlElement && hChild.name.local == 'r') {
        runs.add(RunParser.parse(
          hChild,
          hyperlink: hyperlinkUrl,
          bookmarkRef: bookmarkRef,
        ));
      }
    }
  }

  static XmlElement? _findChild(XmlElement? parent, String localName) {
    if (parent == null) return null;
    for (final child in parent.children) {
      if (child is XmlElement && child.name.local == localName) {
        return child;
      }
    }
    return null;
  }

  static String? _getAttr(XmlElement? element, String name) {
    if (element == null) return null;
    return element.getAttribute('w:$name') ?? element.getAttribute(name);
  }
}

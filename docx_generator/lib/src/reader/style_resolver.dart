import 'package:xml/xml.dart';

import '../models/models.dart';

/// Resolves style IDs and numbering IDs to [DocxParagraphStyle] values.
///
/// Parses `word/styles.xml` and `word/numbering.xml` to build a mapping
/// from DOCX internal identifiers to the library's enum values.
class StyleResolver {
  StyleResolver._();

  /// Known style ID → [DocxParagraphStyle] mapping.
  ///
  /// Covers both docs_gee-generated styleIds and common Word styleIds.
  static const _knownStyles = <String, DocxParagraphStyle>{
    'Normal': DocxParagraphStyle.normal,
    'Heading1': DocxParagraphStyle.heading1,
    'Heading2': DocxParagraphStyle.heading2,
    'Heading3': DocxParagraphStyle.heading3,
    'Heading4': DocxParagraphStyle.heading4,
    'Subtitle': DocxParagraphStyle.subtitle,
    'Caption': DocxParagraphStyle.caption,
    'Quote': DocxParagraphStyle.quote,
    'CodeBlock': DocxParagraphStyle.codeBlock,
    'Footnote': DocxParagraphStyle.footnote,
    'ListBullet': DocxParagraphStyle.listBullet,
    'ListDash': DocxParagraphStyle.listDash,
    'ListNumber': DocxParagraphStyle.listNumber,
    'ListNumberAlpha': DocxParagraphStyle.listNumberAlpha,
    'ListNumberRoman': DocxParagraphStyle.listNumberRoman,
  };

  /// docs_gee numbering: numId → [DocxParagraphStyle].
  static const _numIdToStyle = <int, DocxParagraphStyle>{
    1: DocxParagraphStyle.listBullet,
    2: DocxParagraphStyle.listNumber,
    3: DocxParagraphStyle.listDash,
    4: DocxParagraphStyle.listNumberAlpha,
    5: DocxParagraphStyle.listNumberRoman,
  };

  /// Resolves a `<w:pStyle w:val="..."/>` value to [DocxParagraphStyle].
  static DocxParagraphStyle resolveStyleId(String? styleId) {
    if (styleId == null) return DocxParagraphStyle.normal;
    return _knownStyles[styleId] ?? DocxParagraphStyle.normal;
  }

  /// Resolves a `<w:numId w:val="..."/>` to [DocxParagraphStyle].
  ///
  /// First checks the direct numId mapping (docs_gee convention).
  /// If unknown, falls back to looking up the abstract numbering definition
  /// in [numberingXml] to determine the list type from `numFmt`.
  static DocxParagraphStyle resolveNumId(
    int numId, {
    String? numberingXml,
  }) {
    final directMatch = _numIdToStyle[numId];
    if (directMatch != null) return directMatch;

    if (numberingXml != null) {
      return _resolveFromNumberingXml(numId, numberingXml);
    }

    return DocxParagraphStyle.listBullet;
  }

  /// Parses numbering.xml to find the numFmt for a given numId.
  static DocxParagraphStyle _resolveFromNumberingXml(
    int numId,
    String numberingXml,
  ) {
    final XmlDocument document;
    try {
      document = XmlDocument.parse(numberingXml);
    } on XmlException {
      return DocxParagraphStyle.listBullet;
    }

    // Find <w:num w:numId="numId"> → get abstractNumId
    int? abstractNumId;
    for (final numElement in document.findAllElements('num',
        namespace: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')) {
      final id = int.tryParse(numElement.getAttribute('numId',
              namespace: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') ??
          '');
      if (id == numId) {
        final abstractRef = numElement
            .findAllElements('abstractNumId',
                namespace: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            .firstOrNull;
        abstractNumId = int.tryParse(abstractRef?.getAttribute('val',
                namespace: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') ??
            '');
        break;
      }
    }

    // Also try without namespace (common in generated XML)
    if (abstractNumId == null) {
      for (final numElement in document.findAllElements('num')) {
        final id =
            int.tryParse(_getWAttr(numElement, 'numId') ?? '');
        if (id == numId) {
          final abstractRef =
              _findWElement(numElement, 'abstractNumId');
          abstractNumId =
              int.tryParse(_getWAttr(abstractRef, 'val') ?? '');
          break;
        }
      }
    }

    if (abstractNumId == null) return DocxParagraphStyle.listBullet;

    // Find <w:abstractNum w:abstractNumId="abstractNumId"> → get numFmt from level 0
    return _getStyleFromAbstractNum(document, abstractNumId);
  }

  static DocxParagraphStyle _getStyleFromAbstractNum(
    XmlDocument document,
    int abstractNumId,
  ) {
    for (final abstractNum in document.findAllElements('abstractNum')) {
      final id = int.tryParse(
          _getWAttr(abstractNum, 'abstractNumId') ?? '');
      if (id != abstractNumId) continue;

      // Get level 0 numFmt
      for (final lvl in abstractNum.findAllElements('lvl')) {
        final ilvl =
            int.tryParse(_getWAttr(lvl, 'ilvl') ?? '');
        if (ilvl != 0) continue;

        final numFmtElement = _findWElement(lvl, 'numFmt');
        final numFmt = _getWAttr(numFmtElement, 'val');

        return _numFmtToStyle(numFmt);
      }
    }

    return DocxParagraphStyle.listBullet;
  }

  /// Maps numFmt values to [DocxParagraphStyle].
  static DocxParagraphStyle _numFmtToStyle(String? numFmt) {
    return switch (numFmt) {
      'bullet' => DocxParagraphStyle.listBullet,
      'decimal' => DocxParagraphStyle.listNumber,
      'lowerLetter' => DocxParagraphStyle.listNumberAlpha,
      'upperRoman' => DocxParagraphStyle.listNumberRoman,
      _ => DocxParagraphStyle.listBullet,
    };
  }

  /// Resolves a `<w:jc w:val="..."/>` value to [DocxAlignment].
  static DocxAlignment resolveAlignment(String? value) {
    return switch (value) {
      'left' || 'start' => DocxAlignment.left,
      'center' => DocxAlignment.center,
      'right' || 'end' => DocxAlignment.right,
      'both' || 'justify' => DocxAlignment.justify,
      _ => DocxAlignment.left,
    };
  }

  /// Resolves a `<w:vAlign w:val="..."/>` value to [DocxVerticalAlignment].
  static DocxVerticalAlignment resolveVerticalAlignment(String? value) {
    return switch (value) {
      'top' => DocxVerticalAlignment.top,
      'center' => DocxVerticalAlignment.center,
      'bottom' => DocxVerticalAlignment.bottom,
      _ => DocxVerticalAlignment.top,
    };
  }

  /// Gets an attribute with `w:` prefix (tries both with and without namespace).
  static String? _getWAttr(XmlElement? element, String name) {
    if (element == null) return null;
    return element.getAttribute(name,
            namespace:
                'http://schemas.openxmlformats.org/wordprocessingml/2006/main') ??
        element.getAttribute('w:$name') ??
        element.getAttribute(name);
  }

  /// Finds a child element with `w:` prefix.
  static XmlElement? _findWElement(XmlElement? parent, String localName) {
    if (parent == null) return null;
    return parent
            .findAllElements(localName,
                namespace:
                    'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            .firstOrNull ??
        parent.findAllElements('w:$localName').firstOrNull;
  }
}

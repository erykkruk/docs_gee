import 'package:xml/xml.dart';

import '../models/models.dart';

/// Parses `<w:r>` elements into [DocxRun] objects.
class RunParser {
  const RunParser._();

  /// Word highlight color name → hex mapping.
  static const _highlightToHex = <String, String>{
    'yellow': 'FFFF00',
    'green': '00FF00',
    'cyan': '00FFFF',
    'magenta': 'FF00FF',
    'blue': '0000FF',
    'red': 'FF0000',
    'darkBlue': '000080',
    'darkCyan': '008080',
    'darkGreen': '008000',
    'darkMagenta': '800080',
    'darkRed': '800000',
    'darkYellow': '808000',
    'darkGray': '808080',
    'lightGray': 'C0C0C0',
    'black': '000000',
    'white': 'FFFFFF',
  };

  /// Parses a `<w:r>` element into a [DocxRun].
  ///
  /// [runElement] is the `<w:r>` XML element.
  /// [hyperlink] is the external URL if this run is inside `<w:hyperlink r:id="...">`.
  /// [bookmarkRef] is the anchor name if inside `<w:hyperlink w:anchor="...">`.
  static DocxRun parse(
    XmlElement runElement, {
    String? hyperlink,
    String? bookmarkRef,
  }) {
    // Check for line break: <w:br/> (without type="page")
    final brElement = _findChild(runElement, 'br');
    if (brElement != null) {
      final brType = _getAttr(brElement, 'type');
      if (brType == null) {
        // Simple line break
        return const DocxRun.lineBreak();
      }
      // page break → handled at paragraph level, not as a run
    }

    // Extract text from all <w:t> and <w:br/> children in order
    final text = _extractText(runElement);

    // Parse run properties <w:rPr>
    final rPr = _findChild(runElement, 'rPr');

    final bold = _hasChild(rPr, 'b');
    final italic = _hasChild(rPr, 'i');
    final underline = _hasChild(rPr, 'u');
    final strikethrough = _hasChild(rPr, 'strike');

    // Color: <w:color w:val="FF0000"/>
    String? color;
    final colorElement = _findChild(rPr, 'color');
    if (colorElement != null) {
      color = _getAttr(colorElement, 'val');
      // Filter out auto-applied hyperlink color
      if (hyperlink != null && color == '0000FF') {
        color = null;
      }
    }

    // Background: <w:highlight w:val="yellow"/> or <w:shd w:fill="FFFF00"/>
    String? backgroundColor;
    final highlightElement = _findChild(rPr, 'highlight');
    if (highlightElement != null) {
      final highlightName = _getAttr(highlightElement, 'val');
      if (highlightName != null) {
        backgroundColor = _highlightToHex[highlightName];
      }
    }
    if (backgroundColor == null) {
      final shdElement = _findChild(rPr, 'shd');
      if (shdElement != null) {
        final fill = _getAttr(shdElement, 'fill');
        if (fill != null && fill != 'auto') {
          backgroundColor = fill;
        }
      }
    }

    // Filter out auto-applied hyperlink underline
    final effectiveUnderline =
        underline && !(hyperlink != null && !_hasExplicitUnderline(rPr));

    return DocxRun(
      text,
      bold: bold,
      italic: italic,
      underline: effectiveUnderline,
      strikethrough: strikethrough,
      color: color,
      backgroundColor: backgroundColor,
      hyperlink: hyperlink,
      bookmarkRef: bookmarkRef,
    );
  }

  /// Extracts concatenated text from `<w:t>` elements within a run,
  /// joining with `\n` where `<w:br/>` elements appear between texts.
  static String _extractText(XmlElement runElement) {
    final buffer = StringBuffer();
    var hasPreviousText = false;

    for (final child in runElement.children) {
      if (child is! XmlElement) continue;
      final localName = child.name.local;

      if (localName == 't') {
        buffer.write(child.innerText);
        hasPreviousText = true;
      } else if (localName == 'br') {
        final brType = _getAttr(child, 'type');
        if (brType == null && hasPreviousText) {
          buffer.write('\n');
        }
      }
    }

    return buffer.toString();
  }

  /// Checks if the underline on a hyperlink run was explicitly set
  /// (not just auto-applied by the hyperlink style).
  ///
  /// docs_gee sets underline + color on hyperlinks automatically.
  /// If a run inside a hyperlink has `<w:u>` but no explicit color override,
  /// it's likely the auto-applied style.
  static bool _hasExplicitUnderline(XmlElement? rPr) {
    if (rPr == null) return false;
    final uElement = _findChild(rPr, 'u');
    if (uElement == null) return false;
    // If there's a color element that's NOT the default hyperlink blue,
    // the underline is likely user-specified
    final colorElement = _findChild(rPr, 'color');
    if (colorElement == null) return false;
    final colorVal = _getAttr(colorElement, 'val');
    return colorVal != null && colorVal != '0000FF';
  }

  /// Finds a direct child element by local name (ignoring namespace prefix).
  static XmlElement? _findChild(XmlElement? parent, String localName) {
    if (parent == null) return null;
    for (final child in parent.children) {
      if (child is XmlElement && child.name.local == localName) {
        return child;
      }
    }
    return null;
  }

  /// Checks if a parent element has a child with the given local name.
  static bool _hasChild(XmlElement? parent, String localName) {
    return _findChild(parent, localName) != null;
  }

  /// Gets an attribute value, trying `w:name` then plain `name`.
  static String? _getAttr(XmlElement element, String name) {
    return element.getAttribute('w:$name') ?? element.getAttribute(name);
  }
}

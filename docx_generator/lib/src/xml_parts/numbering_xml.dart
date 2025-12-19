import 'xml_utils.dart';

/// Generates numbering.xml for DOCX (required for lists).
class NumberingXml {
  NumberingXml._();

  static const int _maxLevels = 9;

  /// Generates the numbering.xml content.
  static String generate() {
    final buffer = StringBuffer();
    buffer.writeln(XmlUtils.xmlDeclaration);
    buffer.writeln('<w:numbering xmlns:w="${XmlUtils.wNamespace}">');

    // Abstract numbering 0: Bullet list (•)
    _writeAbstractNum(buffer, 0, 'bullet', '\u2022', 'Symbol');

    // Abstract numbering 1: Numbered list (1, 2, 3...)
    _writeAbstractNumDecimal(buffer, 1, 'decimal', '%1.');

    // Abstract numbering 2: Dash list (-)
    _writeAbstractNum(buffer, 2, 'bullet', '-', 'Courier New');

    // Abstract numbering 3: Alpha list (a, b, c...)
    _writeAbstractNumDecimal(buffer, 3, 'lowerLetter', '%1)');

    // Abstract numbering 4: Roman list (I, II, III...)
    _writeAbstractNumDecimal(buffer, 4, 'upperRoman', '%1.');

    // Numbering instances
    // numId 1 = bullet, numId 2 = decimal, numId 3 = dash, numId 4 = alpha, numId 5 = roman
    for (int i = 0; i < 5; i++) {
      buffer.writeln('  <w:num w:numId="${i + 1}">');
      buffer.writeln('    <w:abstractNumId w:val="$i"/>');
      buffer.writeln('  </w:num>');
    }

    buffer.writeln('</w:numbering>');
    return buffer.toString();
  }

  /// Writes an abstract numbering definition for bullet-style lists.
  static void _writeAbstractNum(
    StringBuffer buffer,
    int abstractNumId,
    String numFmt,
    String symbol,
    String fontName,
  ) {
    buffer.writeln('  <w:abstractNum w:abstractNumId="$abstractNumId">');

    for (int level = 0; level < _maxLevels; level++) {
      final indent =
          720 + (level * 360); // Each level adds 360 twips (0.25 inch)
      buffer.writeln('    <w:lvl w:ilvl="$level">');
      buffer.writeln('      <w:start w:val="1"/>');
      buffer.writeln('      <w:numFmt w:val="$numFmt"/>');
      buffer.writeln('      <w:lvlText w:val="$symbol"/>');
      buffer.writeln('      <w:lvlJc w:val="left"/>');
      buffer.writeln('      <w:pPr>');
      buffer.writeln('        <w:ind w:left="$indent" w:hanging="360"/>');
      buffer.writeln('      </w:pPr>');
      buffer.writeln('      <w:rPr>');
      buffer.writeln(
          '        <w:rFonts w:ascii="$fontName" w:hAnsi="$fontName" w:hint="default"/>');
      buffer.writeln('      </w:rPr>');
      buffer.writeln('    </w:lvl>');
    }

    buffer.writeln('  </w:abstractNum>');
  }

  /// Writes an abstract numbering definition for numbered-style lists.
  static void _writeAbstractNumDecimal(
    StringBuffer buffer,
    int abstractNumId,
    String numFmt,
    String lvlTextPattern,
  ) {
    buffer.writeln('  <w:abstractNum w:abstractNumId="$abstractNumId">');

    for (int level = 0; level < _maxLevels; level++) {
      final indent = 720 + (level * 360);
      // Replace %1 with actual level reference
      final lvlText = lvlTextPattern.replaceAll('%1', '%${level + 1}');
      buffer.writeln('    <w:lvl w:ilvl="$level">');
      buffer.writeln('      <w:start w:val="1"/>');
      buffer.writeln('      <w:numFmt w:val="$numFmt"/>');
      buffer.writeln('      <w:lvlText w:val="$lvlText"/>');
      buffer.writeln('      <w:lvlJc w:val="left"/>');
      buffer.writeln('      <w:pPr>');
      buffer.writeln('        <w:ind w:left="$indent" w:hanging="360"/>');
      buffer.writeln('      </w:pPr>');
      buffer.writeln('    </w:lvl>');
    }

    buffer.writeln('  </w:abstractNum>');
  }
}

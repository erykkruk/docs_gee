import 'xml_utils.dart';

/// Generates numbering.xml for DOCX (required for lists).
class NumberingXml {
  NumberingXml._();

  /// Generates the numbering.xml content.
  static String generate() {
    final buffer = StringBuffer();
    buffer.writeln(XmlUtils.xmlDeclaration);
    buffer.writeln('<w:numbering xmlns:w="${XmlUtils.wNamespace}">');

    // Abstract numbering for bullet list
    buffer.writeln('  <w:abstractNum w:abstractNumId="0">');
    buffer.writeln('    <w:lvl w:ilvl="0">');
    buffer.writeln('      <w:start w:val="1"/>');
    buffer.writeln('      <w:numFmt w:val="bullet"/>');
    buffer.writeln('      <w:lvlText w:val="\u2022"/>'); // bullet character
    buffer.writeln('      <w:lvlJc w:val="left"/>');
    buffer.writeln('      <w:pPr>');
    buffer.writeln('        <w:ind w:left="720" w:hanging="360"/>');
    buffer.writeln('      </w:pPr>');
    buffer.writeln('      <w:rPr>');
    buffer.writeln(
        '        <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>');
    buffer.writeln('      </w:rPr>');
    buffer.writeln('    </w:lvl>');
    buffer.writeln('  </w:abstractNum>');

    // Abstract numbering for numbered list
    buffer.writeln('  <w:abstractNum w:abstractNumId="1">');
    buffer.writeln('    <w:lvl w:ilvl="0">');
    buffer.writeln('      <w:start w:val="1"/>');
    buffer.writeln('      <w:numFmt w:val="decimal"/>');
    buffer.writeln('      <w:lvlText w:val="%1."/>');
    buffer.writeln('      <w:lvlJc w:val="left"/>');
    buffer.writeln('      <w:pPr>');
    buffer.writeln('        <w:ind w:left="720" w:hanging="360"/>');
    buffer.writeln('      </w:pPr>');
    buffer.writeln('    </w:lvl>');
    buffer.writeln('  </w:abstractNum>');

    // Numbering instances
    buffer.writeln('  <w:num w:numId="1">');
    buffer.writeln('    <w:abstractNumId w:val="0"/>');
    buffer.writeln('  </w:num>');
    buffer.writeln('  <w:num w:numId="2">');
    buffer.writeln('    <w:abstractNumId w:val="1"/>');
    buffer.writeln('  </w:num>');

    buffer.writeln('</w:numbering>');
    return buffer.toString();
  }
}

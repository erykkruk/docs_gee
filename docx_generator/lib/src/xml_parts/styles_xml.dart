import 'xml_utils.dart';

/// Generates styles.xml for DOCX.
class StylesXml {
  StylesXml._();

  /// Generates the styles.xml content.
  ///
  /// [fontName] - default font name (e.g., "Times New Roman").
  /// [fontSize] - default font size in half-points (24 = 12pt).
  static String generate({
    String fontName = 'Times New Roman',
    int fontSize = 24,
  }) {
    final buffer = StringBuffer();
    buffer.writeln(XmlUtils.xmlDeclaration);
    buffer.writeln('<w:styles xmlns:w="${XmlUtils.wNamespace}">');

    // Document defaults
    buffer.writeln('  <w:docDefaults>');
    buffer.writeln('    <w:rPrDefault>');
    buffer.writeln('      <w:rPr>');
    buffer
        .writeln('        <w:rFonts w:ascii="$fontName" w:hAnsi="$fontName"/>');
    buffer.writeln('        <w:sz w:val="$fontSize"/>');
    buffer.writeln('        <w:szCs w:val="$fontSize"/>');
    buffer.writeln('      </w:rPr>');
    buffer.writeln('    </w:rPrDefault>');
    buffer.writeln('    <w:pPrDefault>');
    buffer.writeln('      <w:pPr>');
    buffer.writeln(
        '        <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>');
    buffer.writeln('      </w:pPr>');
    buffer.writeln('    </w:pPrDefault>');
    buffer.writeln('  </w:docDefaults>');

    // Normal style
    buffer.writeln(
        '  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">');
    buffer.writeln('    <w:name w:val="Normal"/>');
    buffer.writeln('    <w:qFormat/>');
    buffer.writeln('  </w:style>');

    // Heading 1
    buffer.writeln('  <w:style w:type="paragraph" w:styleId="Heading1">');
    buffer.writeln('    <w:name w:val="heading 1"/>');
    buffer.writeln('    <w:basedOn w:val="Normal"/>');
    buffer.writeln('    <w:next w:val="Normal"/>');
    buffer.writeln('    <w:qFormat/>');
    buffer.writeln('    <w:pPr>');
    buffer.writeln('      <w:keepNext/>');
    buffer.writeln('      <w:spacing w:before="480" w:after="240"/>');
    buffer.writeln('      <w:outlineLvl w:val="0"/>');
    buffer.writeln('    </w:pPr>');
    buffer.writeln('    <w:rPr>');
    buffer.writeln('      <w:b/>');
    buffer.writeln('      <w:sz w:val="48"/>');
    buffer.writeln('      <w:szCs w:val="48"/>');
    buffer.writeln('    </w:rPr>');
    buffer.writeln('  </w:style>');

    // Heading 2
    buffer.writeln('  <w:style w:type="paragraph" w:styleId="Heading2">');
    buffer.writeln('    <w:name w:val="heading 2"/>');
    buffer.writeln('    <w:basedOn w:val="Normal"/>');
    buffer.writeln('    <w:next w:val="Normal"/>');
    buffer.writeln('    <w:qFormat/>');
    buffer.writeln('    <w:pPr>');
    buffer.writeln('      <w:keepNext/>');
    buffer.writeln('      <w:spacing w:before="360" w:after="200"/>');
    buffer.writeln('      <w:outlineLvl w:val="1"/>');
    buffer.writeln('    </w:pPr>');
    buffer.writeln('    <w:rPr>');
    buffer.writeln('      <w:b/>');
    buffer.writeln('      <w:sz w:val="36"/>');
    buffer.writeln('      <w:szCs w:val="36"/>');
    buffer.writeln('    </w:rPr>');
    buffer.writeln('  </w:style>');

    // Heading 3
    buffer.writeln('  <w:style w:type="paragraph" w:styleId="Heading3">');
    buffer.writeln('    <w:name w:val="heading 3"/>');
    buffer.writeln('    <w:basedOn w:val="Normal"/>');
    buffer.writeln('    <w:next w:val="Normal"/>');
    buffer.writeln('    <w:qFormat/>');
    buffer.writeln('    <w:pPr>');
    buffer.writeln('      <w:keepNext/>');
    buffer.writeln('      <w:spacing w:before="280" w:after="160"/>');
    buffer.writeln('      <w:outlineLvl w:val="2"/>');
    buffer.writeln('    </w:pPr>');
    buffer.writeln('    <w:rPr>');
    buffer.writeln('      <w:b/>');
    buffer.writeln('      <w:sz w:val="28"/>');
    buffer.writeln('      <w:szCs w:val="28"/>');
    buffer.writeln('    </w:rPr>');
    buffer.writeln('  </w:style>');

    // Heading 4
    buffer.writeln('  <w:style w:type="paragraph" w:styleId="Heading4">');
    buffer.writeln('    <w:name w:val="heading 4"/>');
    buffer.writeln('    <w:basedOn w:val="Normal"/>');
    buffer.writeln('    <w:next w:val="Normal"/>');
    buffer.writeln('    <w:qFormat/>');
    buffer.writeln('    <w:pPr>');
    buffer.writeln('      <w:keepNext/>');
    buffer.writeln('      <w:spacing w:before="200" w:after="120"/>');
    buffer.writeln('      <w:outlineLvl w:val="3"/>');
    buffer.writeln('    </w:pPr>');
    buffer.writeln('    <w:rPr>');
    buffer.writeln('      <w:b/>');
    buffer.writeln('      <w:sz w:val="26"/>');
    buffer.writeln('      <w:szCs w:val="26"/>');
    buffer.writeln('    </w:rPr>');
    buffer.writeln('  </w:style>');

    // Subtitle
    buffer.writeln('  <w:style w:type="paragraph" w:styleId="Subtitle">');
    buffer.writeln('    <w:name w:val="Subtitle"/>');
    buffer.writeln('    <w:basedOn w:val="Normal"/>');
    buffer.writeln('    <w:next w:val="Normal"/>');
    buffer.writeln('    <w:qFormat/>');
    buffer.writeln('    <w:pPr>');
    buffer.writeln('      <w:spacing w:after="200"/>');
    buffer.writeln('    </w:pPr>');
    buffer.writeln('    <w:rPr>');
    buffer.writeln('      <w:i/>');
    buffer.writeln('      <w:color w:val="5A5A5A"/>');
    buffer.writeln('      <w:sz w:val="28"/>');
    buffer.writeln('      <w:szCs w:val="28"/>');
    buffer.writeln('    </w:rPr>');
    buffer.writeln('  </w:style>');

    // Caption
    buffer.writeln('  <w:style w:type="paragraph" w:styleId="Caption">');
    buffer.writeln('    <w:name w:val="Caption"/>');
    buffer.writeln('    <w:basedOn w:val="Normal"/>');
    buffer.writeln('    <w:qFormat/>');
    buffer.writeln('    <w:pPr>');
    buffer.writeln('      <w:spacing w:before="120" w:after="120"/>');
    buffer.writeln('    </w:pPr>');
    buffer.writeln('    <w:rPr>');
    buffer.writeln('      <w:color w:val="595959"/>');
    buffer.writeln('      <w:sz w:val="20"/>');
    buffer.writeln('      <w:szCs w:val="20"/>');
    buffer.writeln('    </w:rPr>');
    buffer.writeln('  </w:style>');

    // Quote
    buffer.writeln('  <w:style w:type="paragraph" w:styleId="Quote">');
    buffer.writeln('    <w:name w:val="Quote"/>');
    buffer.writeln('    <w:basedOn w:val="Normal"/>');
    buffer.writeln('    <w:next w:val="Normal"/>');
    buffer.writeln('    <w:qFormat/>');
    buffer.writeln('    <w:pPr>');
    buffer.writeln('      <w:spacing w:before="200" w:after="200"/>');
    buffer.writeln('      <w:ind w:left="720" w:right="720"/>');
    buffer.writeln('    </w:pPr>');
    buffer.writeln('    <w:rPr>');
    buffer.writeln('      <w:i/>');
    buffer.writeln('      <w:color w:val="404040"/>');
    buffer.writeln('    </w:rPr>');
    buffer.writeln('  </w:style>');

    // Code Block (monospace font with background)
    buffer.writeln('  <w:style w:type="paragraph" w:styleId="CodeBlock">');
    buffer.writeln('    <w:name w:val="Code Block"/>');
    buffer.writeln('    <w:basedOn w:val="Normal"/>');
    buffer.writeln('    <w:qFormat/>');
    buffer.writeln('    <w:pPr>');
    buffer.writeln('      <w:spacing w:before="120" w:after="120"/>');
    buffer.writeln('      <w:ind w:left="360" w:right="360"/>');
    buffer.writeln('      <w:shd w:val="clear" w:fill="F5F5F5"/>');
    buffer.writeln('    </w:pPr>');
    buffer.writeln('    <w:rPr>');
    buffer.writeln('      <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/>');
    buffer.writeln('      <w:sz w:val="20"/>');
    buffer.writeln('      <w:szCs w:val="20"/>');
    buffer.writeln('    </w:rPr>');
    buffer.writeln('  </w:style>');

    // Footnote Text (smaller font)
    buffer.writeln('  <w:style w:type="paragraph" w:styleId="Footnote">');
    buffer.writeln('    <w:name w:val="Footnote Text"/>');
    buffer.writeln('    <w:basedOn w:val="Normal"/>');
    buffer.writeln('    <w:qFormat/>');
    buffer.writeln('    <w:pPr>');
    buffer.writeln('      <w:spacing w:after="60"/>');
    buffer.writeln('    </w:pPr>');
    buffer.writeln('    <w:rPr>');
    buffer.writeln('      <w:sz w:val="18"/>');
    buffer.writeln('      <w:szCs w:val="18"/>');
    buffer.writeln('      <w:color w:val="666666"/>');
    buffer.writeln('    </w:rPr>');
    buffer.writeln('  </w:style>');

    // List Bullet (•)
    buffer.writeln('  <w:style w:type="paragraph" w:styleId="ListBullet">');
    buffer.writeln('    <w:name w:val="List Bullet"/>');
    buffer.writeln('    <w:basedOn w:val="Normal"/>');
    buffer.writeln('    <w:qFormat/>');
    buffer.writeln('    <w:pPr>');
    buffer.writeln('      <w:numPr>');
    buffer.writeln('        <w:numId w:val="1"/>');
    buffer.writeln('      </w:numPr>');
    buffer.writeln('    </w:pPr>');
    buffer.writeln('  </w:style>');

    // List Number (1, 2, 3...)
    buffer.writeln('  <w:style w:type="paragraph" w:styleId="ListNumber">');
    buffer.writeln('    <w:name w:val="List Number"/>');
    buffer.writeln('    <w:basedOn w:val="Normal"/>');
    buffer.writeln('    <w:qFormat/>');
    buffer.writeln('    <w:pPr>');
    buffer.writeln('      <w:numPr>');
    buffer.writeln('        <w:numId w:val="2"/>');
    buffer.writeln('      </w:numPr>');
    buffer.writeln('    </w:pPr>');
    buffer.writeln('  </w:style>');

    // List Dash (-)
    buffer.writeln('  <w:style w:type="paragraph" w:styleId="ListDash">');
    buffer.writeln('    <w:name w:val="List Dash"/>');
    buffer.writeln('    <w:basedOn w:val="Normal"/>');
    buffer.writeln('    <w:qFormat/>');
    buffer.writeln('    <w:pPr>');
    buffer.writeln('      <w:numPr>');
    buffer.writeln('        <w:numId w:val="3"/>');
    buffer.writeln('      </w:numPr>');
    buffer.writeln('    </w:pPr>');
    buffer.writeln('  </w:style>');

    // List Number Alpha (a, b, c...)
    buffer
        .writeln('  <w:style w:type="paragraph" w:styleId="ListNumberAlpha">');
    buffer.writeln('    <w:name w:val="List Number Alpha"/>');
    buffer.writeln('    <w:basedOn w:val="Normal"/>');
    buffer.writeln('    <w:qFormat/>');
    buffer.writeln('    <w:pPr>');
    buffer.writeln('      <w:numPr>');
    buffer.writeln('        <w:numId w:val="4"/>');
    buffer.writeln('      </w:numPr>');
    buffer.writeln('    </w:pPr>');
    buffer.writeln('  </w:style>');

    // List Number Roman (I, II, III...)
    buffer
        .writeln('  <w:style w:type="paragraph" w:styleId="ListNumberRoman">');
    buffer.writeln('    <w:name w:val="List Number Roman"/>');
    buffer.writeln('    <w:basedOn w:val="Normal"/>');
    buffer.writeln('    <w:qFormat/>');
    buffer.writeln('    <w:pPr>');
    buffer.writeln('      <w:numPr>');
    buffer.writeln('        <w:numId w:val="5"/>');
    buffer.writeln('      </w:numPr>');
    buffer.writeln('    </w:pPr>');
    buffer.writeln('  </w:style>');

    buffer.writeln('</w:styles>');
    return buffer.toString();
  }
}

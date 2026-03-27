import 'package:docs_gee/docs_gee.dart';
import 'package:docs_gee/src/reader/table_parser.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

void main() {
  XmlElement _parseTable(String xml) {
    return XmlDocument.parse(
            '<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '$xml</root>')
        .rootElement
        .children
        .whereType<XmlElement>()
        .first;
  }

  group('TableParser - basic structure', () {
    test('parses simple 2x2 table', () {
      final tbl = _parseTable('''
<w:tbl>
  <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>
  </w:tr>
</w:tbl>''');

      final table = TableParser.parse(tbl, {});
      expect(table.rowCount, 2);
      expect(table.rows[0].cells, hasLength(2));
      expect(table.rows[0].cells[0].paragraphs.first.plainText, 'A');
      expect(table.rows[1].cells[1].paragraphs.first.plainText, 'D');
    });
  });

  group('TableParser - borders', () {
    test('parses table borders', () {
      final tbl = _parseTable('''
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="0" w:type="auto"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:tblBorders>
  </w:tblPr>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr><w:p/></w:tc>
  </w:tr>
</w:tbl>''');

      final table = TableParser.parse(tbl, {});
      expect(table.borders.top, isNotNull);
      expect(table.borders.top!.style, DocxBorderStyle.single);
      expect(table.borders.top!.size, 4);
      expect(table.borders.top!.color, '000000');
      expect(table.borders.insideH, isNotNull);
      expect(table.borders.insideV, isNotNull);
    });

    test('parses nil borders as null', () {
      final tbl = _parseTable('''
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="0" w:type="auto"/>
    <w:tblBorders>
      <w:top w:val="nil"/>
      <w:bottom w:val="nil"/>
      <w:left w:val="nil"/>
      <w:right w:val="nil"/>
      <w:insideH w:val="nil"/>
      <w:insideV w:val="nil"/>
    </w:tblBorders>
  </w:tblPr>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr><w:p/></w:tc>
  </w:tr>
</w:tbl>''');

      final table = TableParser.parse(tbl, {});
      expect(table.borders.top, isNull);
      expect(table.borders.bottom, isNull);
    });

    test('parses dashed border style', () {
      final tbl = _parseTable('''
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="0" w:type="auto"/>
    <w:tblBorders>
      <w:top w:val="dashed" w:sz="8" w:space="0" w:color="FF0000"/>
      <w:bottom w:val="nil"/>
      <w:left w:val="nil"/>
      <w:right w:val="nil"/>
      <w:insideH w:val="nil"/>
      <w:insideV w:val="nil"/>
    </w:tblBorders>
  </w:tblPr>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr><w:p/></w:tc>
  </w:tr>
</w:tbl>''');

      final table = TableParser.parse(tbl, {});
      expect(table.borders.top!.style, DocxBorderStyle.dashed);
      expect(table.borders.top!.size, 8);
      expect(table.borders.top!.color, 'FF0000');
    });
  });

  group('TableParser - colspan', () {
    test('parses gridSpan as colSpan', () {
      final tbl = _parseTable('''
<w:tbl>
  <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w="9360" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr>
      <w:p><w:r><w:t>Spanning</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>''');

      final table = TableParser.parse(tbl, {});
      expect(table.rows[0].cells[0].colSpan, 2);
    });
  });

  group('TableParser - rowspan (vMerge)', () {
    test('parses vMerge restart and continuation', () {
      final tbl = _parseTable('''
<w:tbl>
  <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w="4680" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>
      <w:p><w:r><w:t>Merged</w:t></w:r></w:p>
    </w:tc>
    <w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w="4680" w:type="dxa"/><w:vMerge/></w:tcPr>
      <w:p/>
    </w:tc>
    <w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>
  </w:tr>
</w:tbl>''');

      final table = TableParser.parse(tbl, {});
      expect(table.rows[0].cells[0].rowSpan, 2);
      expect(table.rows[0].cells[0].isMergedContinuation, isFalse);
      expect(table.rows[1].cells[0].isMergedContinuation, isTrue);
    });
  });

  group('TableParser - cell properties', () {
    test('parses background color', () {
      final tbl = _parseTable('''
<w:tbl>
  <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
  <w:tr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="9360" w:type="dxa"/>
        <w:shd w:val="clear" w:fill="E0E0E0"/>
      </w:tcPr>
      <w:p><w:r><w:t>Colored</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>''');

      final table = TableParser.parse(tbl, {});
      expect(table.rows[0].cells[0].backgroundColor, 'E0E0E0');
    });

    test('parses vertical alignment', () {
      final tbl = _parseTable('''
<w:tbl>
  <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
  <w:tr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="9360" w:type="dxa"/>
        <w:vAlign w:val="center"/>
      </w:tcPr>
      <w:p><w:r><w:t>Centered</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>''');

      final table = TableParser.parse(tbl, {});
      expect(table.rows[0].cells[0].verticalAlignment,
          DocxVerticalAlignment.center);
    });

    test('parses cell borders', () {
      final tbl = _parseTable('''
<w:tbl>
  <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
  <w:tr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="9360" w:type="dxa"/>
        <w:tcBorders>
          <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
          <w:bottom w:val="single" w:sz="8" w:space="0" w:color="FF0000"/>
          <w:left w:val="nil"/>
          <w:right w:val="nil"/>
        </w:tcBorders>
      </w:tcPr>
      <w:p><w:r><w:t>Bordered</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>''');

      final table = TableParser.parse(tbl, {});
      final borders = table.rows[0].cells[0].borders!;
      expect(borders.top, isNotNull);
      expect(borders.top!.size, 4);
      expect(borders.bottom, isNotNull);
      expect(borders.bottom!.color, 'FF0000');
      expect(borders.left, isNull);
      expect(borders.right, isNull);
    });
  });
}

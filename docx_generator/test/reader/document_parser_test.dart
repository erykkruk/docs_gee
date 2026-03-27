import 'package:docs_gee/docs_gee.dart';
import 'package:docs_gee/src/reader/document_parser.dart';
import 'package:test/test.dart';

void main() {
  const _ns =
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

  String _wrapBody(String body) =>
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<w:document $_ns><w:body>$body</w:body></w:document>';

  group('DocumentParser - paragraphs', () {
    test('parses plain paragraph', () {
      final xml = _wrapBody('<w:p><w:r><w:t>Hello</w:t></w:r></w:p>');
      final content = DocumentParser.parse(xml, {});

      expect(content, hasLength(1));
      expect(content.first, isA<DocxParagraph>());
      expect((content.first as DocxParagraph).plainText, 'Hello');
    });

    test('parses paragraph with style', () {
      final xml = _wrapBody(
          '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>');
      final content = DocumentParser.parse(xml, {});

      final p = content.first as DocxParagraph;
      expect(p.style, DocxParagraphStyle.heading1);
    });

    test('parses paragraph with alignment', () {
      final xml = _wrapBody(
          '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Centered</w:t></w:r></w:p>');
      final content = DocumentParser.parse(xml, {});

      final p = content.first as DocxParagraph;
      expect(p.alignment, DocxAlignment.center);
    });

    test('parses bookmark name', () {
      final xml = _wrapBody(
          '<w:p><w:bookmarkStart w:id="0" w:name="section1"/>'
          '<w:r><w:t>Section</w:t></w:r>'
          '<w:bookmarkEnd w:id="0"/></w:p>');
      final content = DocumentParser.parse(xml, {});

      final p = content.first as DocxParagraph;
      expect(p.bookmarkName, 'section1');
    });
  });

  group('DocumentParser - lists', () {
    test('parses list with numPr', () {
      final xml = _wrapBody(
          '<w:p><w:pPr><w:pStyle w:val="ListBullet"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
          '<w:r><w:t>Item</w:t></w:r></w:p>');
      final content = DocumentParser.parse(xml, {});

      final p = content.first as DocxParagraph;
      expect(p.style, DocxParagraphStyle.listBullet);
      expect(p.indentLevel, 0);
    });

    test('parses nested list with indent level', () {
      final xml = _wrapBody(
          '<w:p><w:pPr><w:pStyle w:val="ListBullet"/><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr></w:pPr>'
          '<w:r><w:t>Nested</w:t></w:r></w:p>');
      final content = DocumentParser.parse(xml, {});

      final p = content.first as DocxParagraph;
      expect(p.indentLevel, 2);
    });
  });

  group('DocumentParser - page breaks', () {
    test('detects page-break-only paragraph and applies to next', () {
      final xml = _wrapBody(
          '<w:p><w:r><w:t>Before</w:t></w:r></w:p>'
          '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
          '<w:p><w:r><w:t>After</w:t></w:r></w:p>');
      final content = DocumentParser.parse(xml, {});

      expect(content, hasLength(2));
      final first = content[0] as DocxParagraph;
      final second = content[1] as DocxParagraph;
      expect(first.pageBreakBefore, isFalse);
      expect(second.pageBreakBefore, isTrue);
      expect(second.plainText, 'After');
    });
  });

  group('DocumentParser - hyperlinks', () {
    test('parses external hyperlink', () {
      final xml = _wrapBody(
          '<w:p><w:hyperlink r:id="rId100"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p>');
      final rels = {'rId100': 'https://example.com'};
      final content = DocumentParser.parse(xml, rels);

      final p = content.first as DocxParagraph;
      expect(p.runs.first.hyperlink, 'https://example.com');
    });

    test('parses internal bookmark hyperlink', () {
      final xml = _wrapBody(
          '<w:p><w:hyperlink w:anchor="target"><w:r><w:t>Go</w:t></w:r></w:hyperlink></w:p>');
      final content = DocumentParser.parse(xml, {});

      final p = content.first as DocxParagraph;
      expect(p.runs.first.bookmarkRef, 'target');
    });
  });

  group('DocumentParser - mixed content', () {
    test('preserves paragraph-table order', () {
      final xml = _wrapBody(
          '<w:p><w:r><w:t>Before</w:t></w:r></w:p>'
          '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
          '<w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>'
          '</w:tbl>'
          '<w:p><w:r><w:t>After</w:t></w:r></w:p>');
      final content = DocumentParser.parse(xml, {});

      expect(content, hasLength(3));
      expect(content[0], isA<DocxParagraph>());
      expect(content[1], isA<DocxTable>());
      expect(content[2], isA<DocxParagraph>());
    });
  });

  group('DocumentParser - error handling', () {
    test('throws InvalidDocxXmlException for malformed XML', () {
      expect(
        () => DocumentParser.parse('not xml', {}),
        throwsA(isA<InvalidDocxXmlException>()),
      );
    });

    test('throws InvalidDocxXmlException for missing body', () {
      const xml =
          '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:document>';
      expect(
        () => DocumentParser.parse(xml, {}),
        throwsA(isA<InvalidDocxXmlException>()),
      );
    });
  });

  group('DocumentParser - skips sectPr', () {
    test('ignores section properties', () {
      final xml = _wrapBody(
          '<w:p><w:r><w:t>Text</w:t></w:r></w:p>'
          '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>');
      final content = DocumentParser.parse(xml, {});

      expect(content, hasLength(1));
      expect(content.first, isA<DocxParagraph>());
    });
  });
}

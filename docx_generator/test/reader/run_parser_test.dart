import 'package:docs_gee/src/reader/run_parser.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

void main() {
  XmlElement _parseRun(String xml) {
    // Parse XML with namespace, then find first w:r element
    final doc = XmlDocument.parse(
        '<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">$xml</root>');
    return doc.rootElement.children.whereType<XmlElement>().first;
  }

  group('RunParser - text extraction', () {
    test('extracts plain text', () {
      final run = _parseRun('<w:r><w:t>Hello</w:t></w:r>');
      final result = RunParser.parse(run);
      expect(result.text, 'Hello');
    });

    test('extracts text with spaces preserved', () {
      final run =
          _parseRun('<w:r><w:t xml:space="preserve"> Hello </w:t></w:r>');
      final result = RunParser.parse(run);
      expect(result.text, ' Hello ');
    });

    test('returns empty text for empty run', () {
      final run = _parseRun('<w:r></w:r>');
      final result = RunParser.parse(run);
      expect(result.text, '');
    });
  });

  group('RunParser - formatting', () {
    test('parses bold', () {
      final run = _parseRun('<w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>');
      final result = RunParser.parse(run);
      expect(result.bold, isTrue);
      expect(result.italic, isFalse);
    });

    test('parses italic', () {
      final run = _parseRun('<w:r><w:rPr><w:i/></w:rPr><w:t>Italic</w:t></w:r>');
      final result = RunParser.parse(run);
      expect(result.italic, isTrue);
    });

    test('parses underline', () {
      final run = _parseRun(
          '<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>Under</w:t></w:r>');
      final result = RunParser.parse(run);
      expect(result.underline, isTrue);
    });

    test('parses strikethrough', () {
      final run = _parseRun(
          '<w:r><w:rPr><w:strike/></w:rPr><w:t>Strike</w:t></w:r>');
      final result = RunParser.parse(run);
      expect(result.strikethrough, isTrue);
    });

    test('parses color', () {
      final run = _parseRun(
          '<w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>Red</w:t></w:r>');
      final result = RunParser.parse(run);
      expect(result.color, 'FF0000');
    });

    test('parses highlight as background color', () {
      final run = _parseRun(
          '<w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>Hi</w:t></w:r>');
      final result = RunParser.parse(run);
      expect(result.backgroundColor, 'FFFF00');
    });

    test('parses shd fill as background color', () {
      final run = _parseRun(
          '<w:r><w:rPr><w:shd w:val="clear" w:fill="AABBCC"/></w:rPr><w:t>Bg</w:t></w:r>');
      final result = RunParser.parse(run);
      expect(result.backgroundColor, 'AABBCC');
    });

    test('parses combined formatting', () {
      final run = _parseRun(
          '<w:r><w:rPr><w:b/><w:i/><w:u w:val="single"/><w:strike/></w:rPr><w:t>All</w:t></w:r>');
      final result = RunParser.parse(run);
      expect(result.bold, isTrue);
      expect(result.italic, isTrue);
      expect(result.underline, isTrue);
      expect(result.strikethrough, isTrue);
    });
  });

  group('RunParser - line breaks', () {
    test('parses simple line break', () {
      final run = _parseRun('<w:r><w:br/></w:r>');
      final result = RunParser.parse(run);
      expect(result.isLineBreak, isTrue);
    });
  });

  group('RunParser - hyperlinks', () {
    test('passes external hyperlink through', () {
      final run = _parseRun('<w:r><w:t>Link</w:t></w:r>');
      final result =
          RunParser.parse(run, hyperlink: 'https://example.com');
      expect(result.hyperlink, 'https://example.com');
      expect(result.text, 'Link');
    });

    test('passes bookmark reference through', () {
      final run = _parseRun('<w:r><w:t>Ref</w:t></w:r>');
      final result = RunParser.parse(run, bookmarkRef: 'target');
      expect(result.bookmarkRef, 'target');
    });

    test('filters auto-applied hyperlink blue color', () {
      final run = _parseRun(
          '<w:r><w:rPr><w:color w:val="0000FF"/></w:rPr><w:t>Link</w:t></w:r>');
      final result =
          RunParser.parse(run, hyperlink: 'https://example.com');
      expect(result.color, isNull);
    });
  });
}

import 'package:docs_gee/src/docx_reader_exception.dart';
import 'package:docs_gee/src/reader/relationship_resolver.dart';
import 'package:test/test.dart';

void main() {
  group('RelationshipResolver', () {
    test('extracts hyperlink relationships', () {
      const xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
  <Relationship Id="rId101" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://dart.dev" TargetMode="External"/>
</Relationships>''';

      final result = RelationshipResolver.resolve(xml);

      expect(result, hasLength(2));
      expect(result['rId100'], 'https://example.com');
      expect(result['rId101'], 'https://dart.dev');
    });

    test('returns empty map when no hyperlinks', () {
      const xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>''';

      final result = RelationshipResolver.resolve(xml);
      expect(result, isEmpty);
    });

    test('ignores non-hyperlink relationships', () {
      const xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>''';

      final result = RelationshipResolver.resolve(xml);
      expect(result, isEmpty);
    });

    test('throws InvalidDocxXmlException for malformed XML', () {
      expect(
        () => RelationshipResolver.resolve('not xml'),
        throwsA(isA<InvalidDocxXmlException>()),
      );
    });
  });
}

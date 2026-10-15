import 'dart:convert';
import 'dart:typed_data';

import 'package:archive/archive.dart';
import 'package:docs_gee/docs_gee.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

import 'support/docx_archive.dart';

/// Generates a document and reads it straight back.
DocxDocument roundTrip(DocxDocument doc) {
  return const DocxReader().read(DocxGenerator().generate(doc));
}

/// Builds a DOCX whose core.xml is [coreXml], the way Word writes it, with
/// the rest of the package taken from a minimal generated document.
Uint8List docxWithCoreXml(String coreXml) {
  final base = DocxGenerator()
      .generate(DocxDocument()..addParagraph(DocxParagraph.text('Body')));
  final archive = ZipDecoder().decodeBytes(base);
  final coreBytes = utf8.encode(coreXml);
  archive
      .addFile(ArchiveFile('docProps/core.xml', coreBytes.length, coreBytes));
  return Uint8List.fromList(ZipEncoder().encode(archive));
}

void main() {
  group('Core properties - generation', () {
    test('writes title and author to docProps/core.xml', () {
      // Arrange
      final doc = DocxDocument(title: 'Quarterly report', author: 'Jane Doe')
        ..addParagraph(DocxParagraph.text('Body'));

      // Act
      final bytes = DocxGenerator().generate(doc);

      // Assert
      final core = XmlDocument.parse(partText(bytes, 'docProps/core.xml')!);
      expect(core.rootElement.name.local, 'coreProperties');
      expect(core.findAllElements('dc:title').single.innerText,
          'Quarterly report');
      expect(core.findAllElements('dc:creator').single.innerText, 'Jane Doe');
    });

    test('registers the part in content types and package relationships', () {
      // Arrange
      final doc = DocxDocument(title: 'Report')
        ..addParagraph(DocxParagraph.text('Body'));

      // Act
      final bytes = DocxGenerator().generate(doc);

      // Assert
      expect(
        partText(bytes, '[Content_Types].xml'),
        contains('<Override PartName="/docProps/core.xml" '
            'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'),
      );
      final rels = partText(bytes, '_rels/.rels')!;
      expect(rels, contains('metadata/core-properties'));
      expect(rels, contains('Target="docProps/core.xml"'));
    });

    test('escapes XML special characters in metadata', () {
      // Arrange
      final doc = DocxDocument(
        title: 'R&D <draft> "v2"',
        author: "O'Brien",
      )..addParagraph(DocxParagraph.text('Body'));

      // Act
      final bytes = DocxGenerator().generate(doc);

      // Assert
      final core = partText(bytes, 'docProps/core.xml')!;
      expect(() => XmlDocument.parse(core), returnsNormally);
      expect(core, contains('R&amp;D &lt;draft&gt; &quot;v2&quot;'));
    });

    test('writes only the fields that are set', () {
      // Arrange
      final doc = DocxDocument(author: 'Jane Doe')
        ..addParagraph(DocxParagraph.text('Body'));

      // Act
      final bytes = DocxGenerator().generate(doc);

      // Assert
      final core = partText(bytes, 'docProps/core.xml')!;
      expect(core, contains('<dc:creator>Jane Doe</dc:creator>'));
      expect(core, isNot(contains('dc:title')));
    });

    test('adds no part when the document has no metadata', () {
      // Arrange
      final doc = DocxDocument(title: '', author: null)
        ..addParagraph(DocxParagraph.text('Body'));

      // Act
      final bytes = DocxGenerator().generate(doc);

      // Assert
      expect(partNames(bytes), isNot(contains('docProps/core.xml')));
      expect(
          partText(bytes, '_rels/.rels'), isNot(contains('core-properties')));
      expect(partText(bytes, '[Content_Types].xml'),
          isNot(contains('core-properties')));
    });
  });

  group('Core properties - round-trip', () {
    test('title and author survive a round-trip', () {
      // Arrange
      final doc = DocxDocument(title: 'Quarterly report', author: 'Jane Doe')
        ..addParagraph(DocxParagraph.text('Body'));

      // Act
      final result = roundTrip(doc);

      // Assert
      expect(result.title, 'Quarterly report');
      expect(result.author, 'Jane Doe');
    });

    test('special characters come back unescaped', () {
      // Arrange
      final doc = DocxDocument(title: 'R&D <draft>', author: "O'Brien")
        ..addParagraph(DocxParagraph.text('Body'));

      // Act
      final result = roundTrip(doc);

      // Assert
      expect(result.title, 'R&D <draft>');
      expect(result.author, "O'Brien");
    });
  });

  group('Core properties - reading Word output', () {
    test('reads metadata written with the full Word namespace set', () {
      // Arrange
      const coreXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<cp:coreProperties '
          'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
          'xmlns:dc="http://purl.org/dc/elements/1.1/" '
          'xmlns:dcterms="http://purl.org/dc/terms/" '
          'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
          '<dc:title>Budget &amp; Plan</dc:title>'
          '<dc:creator>Jan Kowalski</dc:creator>'
          '<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>'
          '</cp:coreProperties>';

      // Act
      final result = const DocxReader().read(docxWithCoreXml(coreXml));

      // Assert
      expect(result.title, 'Budget & Plan');
      expect(result.author, 'Jan Kowalski');
    });

    test('treats empty metadata elements as unset', () {
      // Arrange
      const coreXml = '<cp:coreProperties '
          'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
          'xmlns:dc="http://purl.org/dc/elements/1.1/">'
          '<dc:title></dc:title><dc:creator/>'
          '</cp:coreProperties>';

      // Act
      final result = const DocxReader().read(docxWithCoreXml(coreXml));

      // Assert
      expect(result.title, isNull);
      expect(result.author, isNull);
    });

    test('ignores a malformed core.xml instead of failing the read', () {
      // Arrange
      const coreXml = '<cp:coreProperties><dc:title>Broken';

      // Act
      final result = const DocxReader().read(docxWithCoreXml(coreXml));

      // Assert
      expect(result.title, isNull);
      expect(result.paragraphs.single.plainText, 'Body');
    });
  });
}

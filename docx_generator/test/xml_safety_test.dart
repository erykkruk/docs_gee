import 'dart:typed_data';

import 'package:docs_gee/docs_gee.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

import 'support/docx_archive.dart';

/// Generates a document and reads it straight back.
DocxDocument roundTrip(DocxDocument doc) {
  return const DocxReader().read(DocxGenerator().generate(doc));
}

/// Parses every XML part of the archive, failing on the first malformed one.
void expectAllPartsWellFormed(Uint8List docx) {
  for (final name in partNames(docx)) {
    if (!name.endsWith('.xml') && !name.endsWith('.rels')) continue;
    expect(
      () => XmlDocument.parse(partText(docx, name)!),
      returnsNormally,
      reason: '$name should be well-formed XML',
    );
  }
}

void main() {
  group('Characters XML 1.0 does not allow', () {
    test('control characters in paragraph text are dropped', () {
      // Arrange
      final doc = DocxDocument()
        ..addParagraph(DocxParagraph.text(
            'Pasted\u000Bfrom\u0000 Word\u000C with a bell\u0007.'));

      // Act
      final bytes = DocxGenerator().generate(doc);

      // Assert
      expectAllPartsWellFormed(bytes);
      expect(partText(bytes, 'word/document.xml'),
          contains('<w:t>Pastedfrom Word with a bell.</w:t>'));
    });

    test('control characters in table cells are dropped', () {
      // Arrange
      final doc = DocxDocument()
        ..addTable(DocxTable.fromHeaders(
          headers: ['Name\u0001'],
          rows: [
            ['Value\u001F'],
          ],
        ));

      // Act
      final bytes = DocxGenerator().generate(doc);

      // Assert
      expectAllPartsWellFormed(bytes);
      final xml = partText(bytes, 'word/document.xml')!;
      expect(xml, contains('Name</w:t>'));
      expect(xml, contains('Value</w:t>'));
    });

    test('control characters in headers and footers are dropped', () {
      // Arrange
      final doc = DocxDocument(
        header: DocxHeaderFooter.text('Head\u0002er'),
        footer: DocxHeaderFooter.text('Foot\u0003er'),
      )..addParagraph(DocxParagraph.text('Body'));

      // Act
      final bytes = DocxGenerator().generate(doc);

      // Assert
      expectAllPartsWellFormed(bytes);
      expect(partText(bytes, 'word/header1.xml'), contains('Header'));
      expect(partText(bytes, 'word/footer1.xml'), contains('Footer'));
    });

    test('noncharacters U+FFFE and U+FFFF are dropped', () {
      // Arrange
      final doc = DocxDocument()..addParagraph(DocxParagraph.text('a￾b￿c'));

      // Act
      final result = roundTrip(doc);

      // Assert
      expect(result.paragraphs.single.plainText, 'abc');
    });

    test('tab and ordinary text are kept untouched', () {
      // Arrange
      final doc = DocxDocument()
        ..addParagraph(DocxParagraph.text('Zażółć\tgęślą & <jaźń>'));

      // Act
      final result = roundTrip(doc);

      // Assert
      expect(result.paragraphs.single.plainText, 'Zażółć\tgęślą & <jaźń>');
    });
  });

  group('Line endings inside a run', () {
    test('Windows line endings become line breaks', () {
      // Arrange
      final doc = DocxDocument()
        ..addParagraph(DocxParagraph.text('one\r\ntwo'));

      // Act
      final bytes = DocxGenerator().generate(doc);

      // Assert
      final xml = partText(bytes, 'word/document.xml')!;
      expect(xml, contains('<w:t>one</w:t><w:br/><w:t>two</w:t>'));
      expect(xml, isNot(contains('\r')));
    });

    test('a bare carriage return becomes a line break', () {
      // Arrange
      final doc = DocxDocument()..addParagraph(DocxParagraph.text('one\rtwo'));

      // Act
      final bytes = DocxGenerator().generate(doc);

      // Assert
      expect(partText(bytes, 'word/document.xml'),
          contains('<w:t>one</w:t><w:br/><w:t>two</w:t>'));
    });

    test('multi-line text survives a round-trip', () {
      // Arrange
      final doc = DocxDocument()
        ..addParagraph(DocxParagraph(runs: const [
          DocxRun('first\r\nsecond\nthird', bold: true),
        ]));

      // Act
      final result = roundTrip(doc);

      // Assert
      final run = result.paragraphs.single.runs.single;
      expect(run.isLineBreak, isFalse);
      expect(run.text, 'first\nsecond\nthird');
      expect(run.bold, isTrue);
    });
  });
}

import 'dart:typed_data';

import 'package:docs_gee/docs_gee.dart';
import 'package:test/test.dart';

import 'support/docx_archive.dart';

/// Point sizes appearing in the PDF content stream, read off the `Tf`
/// operators. The generator writes streams uncompressed, so they can be
/// scanned as text.
Set<String> fontOperators(Uint8List pdf) {
  final text = String.fromCharCodes(pdf);
  return RegExp(r'/F\d+ ([\d.]+) Tf')
      .allMatches(text)
      .map((m) => m.group(1)!.replaceAll(RegExp(r'\.0$'), ''))
      .toSet();
}

/// Generates a document and reads it straight back, which is how the suite
/// checks that a property survives the full write/read cycle.
DocxDocument roundTrip(DocxDocument doc) {
  return const DocxReader().read(DocxGenerator().generate(doc));
}

void main() {
  group('Per-run font size - DOCX', () {
    test('writes the size in half-points', () {
      final doc = DocxDocument()
        ..addParagraph(const DocxParagraph(
          runs: [DocxRun('Large', fontSize: 14)],
        ));

      final document =
          partText(DocxGenerator().generate(doc), 'word/document.xml')!;

      expect(document, contains('<w:sz w:val="28"/>'));
      expect(document, contains('<w:szCs w:val="28"/>'));
    });

    test('writes no size element for a run that does not set one', () {
      final doc = DocxDocument()
        ..addParagraph(const DocxParagraph(runs: [DocxRun('Default')]));

      final document =
          partText(DocxGenerator().generate(doc), 'word/document.xml')!;

      expect(document, isNot(contains('<w:sz ')));
    });

    test('a size makes an otherwise plain run formatted', () {
      expect(const DocxRun('x').hasFormatting, isFalse);
      expect(const DocxRun('x', fontSize: 20).hasFormatting, isTrue);
    });

    test('survives a round-trip through the reader', () {
      final doc = DocxDocument()
        ..addParagraph(const DocxParagraph(
          runs: [
            DocxRun('Small', fontSize: 8),
            DocxRun('Normal'),
            DocxRun('Huge', fontSize: 32),
          ],
        ));

      final runs = roundTrip(doc).paragraphs.first.runs;

      expect(runs[0].fontSize, 8);
      expect(runs[1].fontSize, isNull);
      expect(runs[2].fontSize, 32);
    });

    test('sizes inside table cells round-trip too', () {
      final doc = DocxDocument()
        ..addTable(const DocxTable(
          rows: [
            DocxTableRow(cells: [
              DocxTableCell(
                paragraphs: [
                  DocxParagraph(runs: [DocxRun('Cell', fontSize: 18)]),
                ],
              ),
            ]),
          ],
        ));

      final cell = roundTrip(doc).tables.first.rows.first.cells.first;

      expect(cell.paragraphs.first.runs.first.fontSize, 18);
    });

    test('copyWith carries the size over', () {
      const run = DocxRun('x', fontSize: 11);

      expect(run.copyWith(text: 'y').fontSize, 11);
    });

    test('a field run can set its own size', () {
      final doc = DocxDocument(
        footer: const DocxHeaderFooter(
          paragraphs: [
            DocxParagraph(runs: [DocxRun.pageNumber(fontSize: 9)]),
          ],
        ),
      )..addParagraph(DocxParagraph.text('Body'));

      final footer =
          partText(DocxGenerator().generate(doc), 'word/footer1.xml')!;

      expect(footer, contains('<w:sz w:val="18"/>'));
    });
  });

  group('Per-run font size - PDF', () {
    test('a document with mixed sizes still generates', () {
      final doc = DocxDocument()
        ..addParagraph(const DocxParagraph(
          runs: [
            DocxRun('Small ', fontSize: 6),
            DocxRun('and huge', fontSize: 28),
          ],
        ));

      final bytes = PdfGenerator().generate(doc);

      expect(bytes.length, greaterThan(0));
      expect(String.fromCharCodes(bytes.take(5)), '%PDF-');
    });

    test('the run size reaches the content stream as a Tf operator', () {
      final plain = DocxDocument()
        ..addParagraph(const DocxParagraph(runs: [DocxRun('Sample text')]));
      final sized = DocxDocument()
        ..addParagraph(const DocxParagraph(
          runs: [DocxRun('Sample text', fontSize: 36)],
        ));

      expect(fontOperators(PdfGenerator().generate(plain)), contains('12'));
      expect(fontOperators(PdfGenerator().generate(sized)), contains('36'));
    });

    test('the paragraph style size still applies to runs without one', () {
      final doc = DocxDocument()
        ..addParagraph(const DocxParagraph(
          runs: [DocxRun('Big', fontSize: 30), DocxRun(' and default')],
        ));

      final sizes = fontOperators(PdfGenerator().generate(doc));

      expect(sizes, contains('30'));
      expect(sizes, contains('12'));
    });
  });

  group('Right-to-left support', () {
    test('a right-to-left paragraph writes w:bidi', () {
      final doc = DocxDocument()
        ..addParagraph(const DocxParagraph(
          runs: [DocxRun('שלום', rtl: true)],
          rtl: true,
        ));

      final document =
          partText(DocxGenerator().generate(doc), 'word/document.xml')!;

      expect(document, contains('<w:bidi/>'));
      expect(document, contains('<w:rtl/>'));
    });

    test('w:bidi comes before w:jc so alignment reads in the flipped direction',
        () {
      final doc = DocxDocument()
        ..addParagraph(const DocxParagraph(
          runs: [DocxRun('نص')],
          rtl: true,
          alignment: DocxAlignment.right,
        ));

      final document =
          partText(DocxGenerator().generate(doc), 'word/document.xml')!;

      expect(
          document.indexOf('<w:bidi/>'), lessThan(document.indexOf('<w:jc')));
    });

    test('a left-to-right document writes neither element', () {
      final doc = DocxDocument()
        ..addParagraph(DocxParagraph.text('Plain English'));

      final document =
          partText(DocxGenerator().generate(doc), 'word/document.xml')!;

      expect(document, isNot(contains('<w:bidi/>')));
      expect(document, isNot(contains('<w:rtl/>')));
    });

    test('an rtl run counts as formatted', () {
      expect(const DocxRun('x', rtl: true).hasFormatting, isTrue);
    });

    test('paragraph and run direction survive a round-trip', () {
      final doc = DocxDocument()
        ..addParagraph(const DocxParagraph(
          runs: [DocxRun('مرحبا', rtl: true)],
          rtl: true,
        ));

      final paragraph = roundTrip(doc).paragraphs.first;

      expect(paragraph.rtl, isTrue);
      expect(paragraph.runs.first.rtl, isTrue);
    });

    test('a left-to-right document round-trips as left-to-right', () {
      final doc = DocxDocument()..addParagraph(DocxParagraph.text('Hello'));

      final paragraph = roundTrip(doc).paragraphs.first;

      expect(paragraph.rtl, isFalse);
      expect(paragraph.runs.first.rtl, isFalse);
    });

    test('right-to-left cells round-trip inside tables', () {
      final doc = DocxDocument()
        ..addTable(const DocxTable(
          rows: [
            DocxTableRow(cells: [
              DocxTableCell(
                paragraphs: [
                  DocxParagraph(runs: [DocxRun('עברית', rtl: true)], rtl: true),
                ],
              ),
            ]),
          ],
        ));

      final cell = roundTrip(doc).tables.first.rows.first.cells.first;

      expect(cell.paragraphs.first.rtl, isTrue);
    });
  });

  group('Table cell padding', () {
    test('.points() converts to twips', () {
      const padding = DocxCellPadding.points(top: 6, left: 10);

      expect(padding.top, 120);
      expect(padding.left, 200);
    });

    test('.all() applies one value to every edge', () {
      const padding = DocxCellPadding.all(80);

      expect(
        [padding.top, padding.right, padding.bottom, padding.left],
        everyElement(80),
      );
    });

    test('hasPadding is false only when every edge is zero', () {
      expect(const DocxCellPadding().hasPadding, isFalse);
      expect(const DocxCellPadding(top: 1).hasPadding, isTrue);
    });

    test('value equality compares all four edges', () {
      expect(
        const DocxCellPadding.all(100),
        equals(const DocxCellPadding(
            top: 100, right: 100, bottom: 100, left: 100)),
      );
      expect(
        const DocxCellPadding.all(100),
        isNot(equals(const DocxCellPadding.all(101))),
      );
    });

    test('the Word default matches Word own margins', () {
      expect(DocxCellPadding.wordDefault.left, 108);
      expect(DocxCellPadding.wordDefault.right, 108);
      expect(DocxCellPadding.wordDefault.top, 0);
    });

    test('table-wide padding writes tblCellMar', () {
      final doc = DocxDocument()
        ..addTable(DocxTable.simple(
          [
            ['A', 'B'],
          ],
          cellPadding: const DocxCellPadding.points(top: 4, left: 8),
        ));

      final document =
          partText(DocxGenerator().generate(doc), 'word/document.xml')!;

      expect(document, contains('<w:tblCellMar>'));
      expect(document, contains('<w:top w:w="80" w:type="dxa"/>'));
      expect(document, contains('<w:left w:w="160" w:type="dxa"/>'));
    });

    test('per-cell padding writes tcMar', () {
      final doc = DocxDocument()
        ..addTable(DocxTable(
          rows: [
            DocxTableRow(cells: [
              DocxTableCell.text(
                'Padded',
                padding: const DocxCellPadding.all(150),
              ),
            ]),
          ],
        ));

      final document =
          partText(DocxGenerator().generate(doc), 'word/document.xml')!;

      expect(document, contains('<w:tcMar>'));
      expect(document, contains('w:w="150"'));
    });

    test('a table without padding writes no margin elements', () {
      final doc = DocxDocument()
        ..addTable(DocxTable.simple([
          ['A'],
        ]));

      final document =
          partText(DocxGenerator().generate(doc), 'word/document.xml')!;

      expect(document, isNot(contains('<w:tblCellMar>')));
      expect(document, isNot(contains('<w:tcMar>')));
    });

    test('table padding survives a round-trip', () {
      final doc = DocxDocument()
        ..addTable(DocxTable.simple(
          [
            ['A', 'B'],
          ],
          cellPadding: const DocxCellPadding.points(
              top: 5, bottom: 5, left: 7, right: 7),
        ));

      final table = roundTrip(doc).tables.first;

      expect(table.cellPadding, isNotNull);
      expect(table.cellPadding!.top, 100);
      expect(table.cellPadding!.left, 140);
    });

    test('cell padding survives a round-trip', () {
      final doc = DocxDocument()
        ..addTable(DocxTable(
          rows: [
            DocxTableRow(cells: [
              DocxTableCell.text(
                'Padded',
                padding: const DocxCellPadding.all(90),
              ),
            ]),
          ],
        ));

      final cell = roundTrip(doc).tables.first.rows.first.cells.first;

      expect(cell.padding, const DocxCellPadding.all(90));
    });

    test('an unpadded table round-trips back to null, not to all-zero', () {
      final doc = DocxDocument()
        ..addTable(DocxTable.simple([
          ['A'],
        ]));

      final table = roundTrip(doc).tables.first;

      expect(table.cellPadding, isNull);
      expect(table.rows.first.cells.first.padding, isNull);
    });

    test('padding reaches the PDF generator without breaking output', () {
      final doc = DocxDocument()
        ..addTable(DocxTable.simple(
          [
            ['A', 'B'],
            ['C', 'D'],
          ],
          cellPadding: const DocxCellPadding.points(
              top: 8, bottom: 8, left: 12, right: 12),
        ));

      final bytes = PdfGenerator().generate(doc);

      expect(String.fromCharCodes(bytes.take(5)), '%PDF-');
    });
  });
}

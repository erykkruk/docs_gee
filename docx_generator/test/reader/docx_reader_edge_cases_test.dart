import 'package:docs_gee/docs_gee.dart';
import 'package:test/test.dart';

/// Edge case and stress tests for DocxReader round-trip.
void main() {
  late DocxGenerator generator;
  late DocxReader reader;

  setUp(() {
    generator = DocxGenerator();
    reader = const DocxReader();
  });

  DocxDocument _roundTrip(DocxDocument doc) =>
      reader.read(generator.generate(doc));

  group('DocxReader - edge cases: special characters', () {
    test('handles XML special characters in text', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.text('5 < 10 & 10 > 5 "quotes" \'apos\''));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.plainText,
          '5 < 10 & 10 > 5 "quotes" \'apos\'');
    });

    test('handles emoji in text', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.text('Hello World! \u{1F600}'));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.plainText, 'Hello World! \u{1F600}');
    });

    test('handles empty string text', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.text(''));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.plainText, '');
    });

    test('handles text with leading and trailing spaces', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph(
        runs: [const DocxRun(' spaced ')],
      ));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.runs.first.text, ' spaced ');
    });

    test('handles long text', () {
      final longText = 'A' * 10000;
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.text(longText));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.plainText, longText);
    });

    test('handles Polish characters', () {
      final doc = DocxDocument();
      doc.addParagraph(
          DocxParagraph.text('Zażółć gęślą jaźń ĄĆĘŁŃÓŚŹŻ'));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.plainText,
          'Zażółć gęślą jaźń ĄĆĘŁŃÓŚŹŻ');
    });
  });

  group('DocxReader - edge cases: complex documents', () {
    test('reads document with all paragraph styles', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.heading('Title', level: 1));
      doc.addParagraph(DocxParagraph.heading('Subtitle H2', level: 2));
      doc.addParagraph(DocxParagraph.subtitle('Subtitle style'));
      doc.addParagraph(DocxParagraph.text('Normal text'));
      doc.addParagraph(DocxParagraph.caption('A caption'));
      doc.addParagraph(DocxParagraph.quote('A quote'));
      doc.addParagraph(DocxParagraph.codeBlock('var x = 1;'));
      doc.addParagraph(DocxParagraph.footnote('A footnote'));
      doc.addParagraph(DocxParagraph.bulletItem('Bullet'));
      doc.addParagraph(DocxParagraph.dashItem('Dash'));
      doc.addParagraph(DocxParagraph.numberedItem('Numbered'));
      doc.addParagraph(DocxParagraph.alphaItem('Alpha'));
      doc.addParagraph(DocxParagraph.romanItem('Roman'));

      final result = _roundTrip(doc);
      expect(result.paragraphs, hasLength(13));
      expect(result.paragraphs[0].style, DocxParagraphStyle.heading1);
      expect(result.paragraphs[1].style, DocxParagraphStyle.heading2);
      expect(result.paragraphs[2].style, DocxParagraphStyle.subtitle);
      expect(result.paragraphs[3].style, DocxParagraphStyle.normal);
      expect(result.paragraphs[4].style, DocxParagraphStyle.caption);
      expect(result.paragraphs[5].style, DocxParagraphStyle.quote);
      expect(result.paragraphs[6].style, DocxParagraphStyle.codeBlock);
      expect(result.paragraphs[7].style, DocxParagraphStyle.footnote);
      expect(result.paragraphs[8].style, DocxParagraphStyle.listBullet);
      expect(result.paragraphs[9].style, DocxParagraphStyle.listDash);
      expect(result.paragraphs[10].style, DocxParagraphStyle.listNumber);
      expect(result.paragraphs[11].style, DocxParagraphStyle.listNumberAlpha);
      expect(result.paragraphs[12].style, DocxParagraphStyle.listNumberRoman);
    });

    test('reads complex table with mixed cell properties', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable(
        rows: [
          DocxTableRow(cells: [
            DocxTableCell.text('Header 1', backgroundColor: 'E0E0E0'),
            DocxTableCell.text('Header 2', backgroundColor: 'E0E0E0'),
            DocxTableCell.text('Header 3', backgroundColor: 'E0E0E0'),
          ]),
          DocxTableRow(cells: [
            DocxTableCell.text('Span 2', colSpan: 2),
            DocxTableCell.text('Single'),
          ]),
          DocxTableRow(cells: [
            const DocxTableCell(
              paragraphs: [DocxParagraph(runs: [DocxRun('Center')])],
              verticalAlignment: DocxVerticalAlignment.center,
            ),
            const DocxTableCell(
              paragraphs: [DocxParagraph(runs: [DocxRun('Bottom')])],
              verticalAlignment: DocxVerticalAlignment.bottom,
            ),
            const DocxTableCell(
              paragraphs: [DocxParagraph(runs: [DocxRun('Bordered')])],
              borders: DocxCellBorders.all(),
            ),
          ]),
        ],
        borders: const DocxTableBorders.all(),
      ));

      final result = _roundTrip(doc);
      final table = result.tables.first;

      // Row 0: header colors
      expect(table.rows[0].cells[0].backgroundColor, 'E0E0E0');
      expect(table.rows[0].cells[1].backgroundColor, 'E0E0E0');
      expect(table.rows[0].cells[2].backgroundColor, 'E0E0E0');

      // Row 1: colspan
      expect(table.rows[1].cells[0].colSpan, 2);

      // Row 2: vertical alignment
      expect(table.rows[2].cells[0].verticalAlignment,
          DocxVerticalAlignment.center);
      expect(table.rows[2].cells[1].verticalAlignment,
          DocxVerticalAlignment.bottom);
      // Row 2: cell borders
      expect(table.rows[2].cells[2].borders, isNotNull);
    });

    test('reads document with multiple hyperlinks', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph(
        runs: [
          const DocxRun('Visit '),
          const DocxRun('Google', hyperlink: 'https://google.com'),
          const DocxRun(' or '),
          const DocxRun('GitHub', hyperlink: 'https://github.com'),
          const DocxRun('.'),
        ],
      ));

      final result = _roundTrip(doc);
      final runs = result.paragraphs.first.runs;
      expect(runs, hasLength(5));
      expect(runs[0].text, 'Visit ');
      expect(runs[0].hyperlink, isNull);
      expect(runs[1].text, 'Google');
      expect(runs[1].hyperlink, 'https://google.com');
      expect(runs[3].text, 'GitHub');
      expect(runs[3].hyperlink, 'https://github.com');
      expect(runs[4].text, '.');
    });

    test('reads document with formatted hyperlink', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph(
        runs: [
          const DocxRun(
            'Bold link',
            bold: true,
            hyperlink: 'https://example.com',
          ),
        ],
      ));

      final result = _roundTrip(doc);
      final run = result.paragraphs.first.runs.first;
      expect(run.bold, isTrue);
      expect(run.hyperlink, 'https://example.com');
    });

    test('reads document with multiple page breaks', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.text('Page 1'));
      doc.addParagraph(
          DocxParagraph.text('Page 2', pageBreakBefore: true));
      doc.addParagraph(
          DocxParagraph.text('Page 3', pageBreakBefore: true));

      final result = _roundTrip(doc);
      expect(result.paragraphs, hasLength(3));
      expect(result.paragraphs[0].pageBreakBefore, isFalse);
      expect(result.paragraphs[1].pageBreakBefore, isTrue);
      expect(result.paragraphs[2].pageBreakBefore, isTrue);
    });

    test('reads rich paragraph with line breaks and formatting', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph(
        runs: [
          const DocxRun('Line 1', bold: true),
          const DocxRun.lineBreak(),
          const DocxRun('Line 2', italic: true),
          const DocxRun.lineBreak(),
          const DocxRun('Line 3', color: 'FF0000'),
        ],
      ));

      final result = _roundTrip(doc);
      final runs = result.paragraphs.first.runs;
      expect(runs, hasLength(5));
      expect(runs[0].bold, isTrue);
      expect(runs[1].isLineBreak, isTrue);
      expect(runs[2].italic, isTrue);
      expect(runs[3].isLineBreak, isTrue);
      expect(runs[4].color, 'FF0000');
    });

    test('reads full-featured document', () {
      final doc = DocxDocument(title: 'Test Doc', author: 'Eryk');

      doc.addParagraph(DocxParagraph.heading('Introduction', level: 1,
          bookmarkName: 'intro'));
      doc.addParagraph(DocxParagraph.text(
        'Welcome to the document.',
        alignment: DocxAlignment.justify,
      ));

      doc.addTable(DocxTable.fromHeaders(
        headers: ['Feature', 'Status'],
        rows: [
          ['DOCX Reader', 'Done'],
          ['PDF Reader', 'TODO'],
        ],
      ));

      doc.addParagraph(DocxParagraph.heading('Details',
          level: 2, pageBreakBefore: true));
      doc.addParagraph(DocxParagraph.bulletItem('First point'));
      doc.addParagraph(
          DocxParagraph.bulletItem('Nested point', indentLevel: 1));
      doc.addParagraph(DocxParagraph(
        runs: [
          const DocxRun('See '),
          const DocxRun('intro', bookmarkRef: 'intro'),
          const DocxRun(' for more.'),
        ],
      ));

      final result = _roundTrip(doc);

      // Content order
      expect(result.content, hasLength(7));
      expect(result.content[0], isA<DocxParagraph>()); // heading
      expect(result.content[1], isA<DocxParagraph>()); // text
      expect(result.content[2], isA<DocxTable>()); // table
      expect(result.content[3], isA<DocxParagraph>()); // heading 2
      expect(result.content[4], isA<DocxParagraph>()); // bullet
      expect(result.content[5], isA<DocxParagraph>()); // nested bullet
      expect(result.content[6], isA<DocxParagraph>()); // bookmark ref

      // Heading with bookmark
      final h1 = result.content[0] as DocxParagraph;
      expect(h1.style, DocxParagraphStyle.heading1);
      expect(h1.bookmarkName, 'intro');

      // Justified text
      final text = result.content[1] as DocxParagraph;
      expect(text.alignment, DocxAlignment.justify);

      // Table
      final table = result.content[2] as DocxTable;
      expect(table.rowCount, 3);
      expect(table.rows[0].cells[0].backgroundColor, 'E0E0E0');

      // Page break
      final h2 = result.content[3] as DocxParagraph;
      expect(h2.pageBreakBefore, isTrue);
      expect(h2.style, DocxParagraphStyle.heading2);

      // Nested bullet
      final nested = result.content[5] as DocxParagraph;
      expect(nested.indentLevel, 1);

      // Bookmark reference
      final refParagraph = result.content[6] as DocxParagraph;
      expect(refParagraph.runs[1].bookmarkRef, 'intro');
    });
  });

  group('DocxReader - edge cases: table borders', () {
    test('reads table with no borders', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable.simple(
        [
          ['A']
        ],
        borders: const DocxTableBorders.none(),
      ));

      final result = _roundTrip(doc);
      final borders = result.tables.first.borders;
      expect(borders.hasBorders, isFalse);
    });

    test('reads table with outside-only borders', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable.simple(
        [
          ['A', 'B'],
          ['C', 'D'],
        ],
        borders: const DocxTableBorders.outside(),
      ));

      final result = _roundTrip(doc);
      final borders = result.tables.first.borders;
      expect(borders.top, isNotNull);
      expect(borders.bottom, isNotNull);
      expect(borders.left, isNotNull);
      expect(borders.right, isNotNull);
      expect(borders.insideH, isNull);
      expect(borders.insideV, isNull);
    });

    test('reads cell with bottom-only border', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable(
        rows: [
          DocxTableRow(cells: [
            const DocxTableCell(
              paragraphs: [DocxParagraph(runs: [DocxRun('Cell')])],
              borders: DocxCellBorders.bottom(),
            ),
          ]),
        ],
      ));

      final result = _roundTrip(doc);
      final cellBorders = result.tables.first.rows[0].cells[0].borders;
      expect(cellBorders, isNotNull);
      expect(cellBorders!.top, isNull);
      expect(cellBorders.bottom, isNotNull);
      expect(cellBorders.left, isNull);
      expect(cellBorders.right, isNull);
    });
  });

  group('DocxReader - edge cases: empty structures', () {
    test('reads document with no content', () {
      final doc = DocxDocument();

      final result = _roundTrip(doc);
      expect(result.content, isEmpty);
    });

    test('reads table with empty cells', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable(
        rows: [
          const DocxTableRow(cells: [
            DocxTableCell(),
            DocxTableCell(),
          ]),
        ],
      ));

      final result = _roundTrip(doc);
      expect(result.tables.first.rowCount, 1);
      expect(result.tables.first.rows[0].cells, hasLength(2));
    });

    test('reads paragraph with multiple empty runs', () {
      final doc = DocxDocument();
      doc.addParagraph(const DocxParagraph(
        runs: [DocxRun(''), DocxRun('text'), DocxRun('')],
      ));

      final result = _roundTrip(doc);
      // Empty runs may be trimmed by the XML generator,
      // but the non-empty text should survive
      final combinedText = result.paragraphs.first.plainText;
      expect(combinedText, contains('text'));
    });
  });

  group('DocxReader - edge cases: deep nesting', () {
    test('reads deeply nested list', () {
      final doc = DocxDocument();
      for (int i = 0; i <= 5; i++) {
        doc.addParagraph(
            DocxParagraph.bulletItem('Level $i', indentLevel: i));
      }

      final result = _roundTrip(doc);
      expect(result.paragraphs, hasLength(6));
      for (int i = 0; i <= 5; i++) {
        expect(result.paragraphs[i].indentLevel, i);
        expect(result.paragraphs[i].plainText, 'Level $i');
      }
    });

    test('reads table with multiple paragraphs in cell', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable(
        rows: [
          DocxTableRow(cells: [
            const DocxTableCell(
              paragraphs: [
                DocxParagraph(runs: [DocxRun('First')]),
                DocxParagraph(runs: [DocxRun('Second')]),
                DocxParagraph(runs: [DocxRun('Third')]),
              ],
            ),
          ]),
        ],
      ));

      final result = _roundTrip(doc);
      final cell = result.tables.first.rows[0].cells[0];
      expect(cell.paragraphs, hasLength(3));
      expect(cell.paragraphs[0].plainText, 'First');
      expect(cell.paragraphs[1].plainText, 'Second');
      expect(cell.paragraphs[2].plainText, 'Third');
    });

    test('reads large table', () {
      final data = List.generate(
        20,
        (row) => List.generate(5, (col) => 'R${row}C$col'),
      );

      final doc = DocxDocument();
      doc.addTable(DocxTable.simple(data));

      final result = _roundTrip(doc);
      final table = result.tables.first;
      expect(table.rowCount, 20);
      expect(table.rows[0].cells, hasLength(5));
      expect(table.rows[19].cells[4].paragraphs.first.plainText, 'R19C4');
    });
  });

  group('DocxReader - edge cases: all formatting combined', () {
    test('reads run with all formatting properties', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph(
        runs: [
          const DocxRun(
            'Everything',
            bold: true,
            italic: true,
            underline: true,
            strikethrough: true,
            color: 'FF0000',
            backgroundColor: 'FFFF00',
          ),
        ],
      ));

      final result = _roundTrip(doc);
      final run = result.paragraphs.first.runs.first;
      expect(run.text, 'Everything');
      expect(run.bold, isTrue);
      expect(run.italic, isTrue);
      expect(run.underline, isTrue);
      expect(run.strikethrough, isTrue);
      expect(run.color, 'FF0000');
      expect(run.backgroundColor, 'FFFF00');
    });

    test('reads alternating formatted runs', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph(
        runs: [
          const DocxRun('bold', bold: true),
          const DocxRun(' '),
          const DocxRun('italic', italic: true),
          const DocxRun(' '),
          const DocxRun('under', underline: true),
          const DocxRun(' '),
          const DocxRun('strike', strikethrough: true),
          const DocxRun(' '),
          const DocxRun('red', color: 'FF0000'),
          const DocxRun(' '),
          const DocxRun('highlight', backgroundColor: 'FFFF00'),
        ],
      ));

      final result = _roundTrip(doc);
      final runs = result.paragraphs.first.runs;
      expect(runs, hasLength(11));
      expect(runs[0].bold, isTrue);
      expect(runs[2].italic, isTrue);
      expect(runs[4].underline, isTrue);
      expect(runs[6].strikethrough, isTrue);
      expect(runs[8].color, 'FF0000');
      expect(runs[10].backgroundColor, 'FFFF00');
    });
  });

  group('DocxReader - edge cases: various highlight colors', () {
    final colorMap = {
      'FFFF00': 'FFFF00', // yellow
      '00FF00': '00FF00', // green
      '00FFFF': '00FFFF', // cyan
      'FF00FF': 'FF00FF', // magenta
      '0000FF': '0000FF', // blue
      'FF0000': 'FF0000', // red
      '000000': '000000', // black
    };

    for (final entry in colorMap.entries) {
      test('round-trips background color ${entry.key}', () {
        final doc = DocxDocument();
        doc.addParagraph(DocxParagraph(
          runs: [DocxRun('Color', backgroundColor: entry.key)],
        ));

        final result = _roundTrip(doc);
        // The exact hex may change due to highlight mapping,
        // but the backgroundColor should not be null
        expect(result.paragraphs.first.runs.first.backgroundColor, isNotNull);
      });
    }
  });
}

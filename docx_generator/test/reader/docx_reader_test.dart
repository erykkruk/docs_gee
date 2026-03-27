import 'dart:typed_data';

import 'package:docs_gee/docs_gee.dart';
import 'package:test/test.dart';

/// Round-trip tests: generate DOCX → read it back → verify model.
void main() {
  late DocxGenerator generator;
  late DocxReader reader;

  setUp(() {
    generator = DocxGenerator();
    reader = const DocxReader();
  });

  Uint8List _generate(DocxDocument doc) => generator.generate(doc);
  DocxDocument _roundTrip(DocxDocument doc) => reader.read(_generate(doc));

  group('DocxReader - error handling', () {
    test('throws DocxReaderException for invalid bytes', () {
      expect(
        () => reader.read(Uint8List.fromList([1, 2, 3, 4])),
        throwsA(isA<DocxReaderException>()),
      );
    });

    test('throws DocxReaderException for empty bytes', () {
      expect(
        () => reader.read(Uint8List(0)),
        throwsA(isA<DocxReaderException>()),
      );
    });
  });

  group('DocxReader - round-trip: plain text', () {
    test('reads single paragraph', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.text('Hello World'));

      final result = _roundTrip(doc);
      final paragraphs = result.paragraphs;

      expect(paragraphs, hasLength(1));
      expect(paragraphs.first.plainText, 'Hello World');
      expect(paragraphs.first.style, DocxParagraphStyle.normal);
    });

    test('reads multiple paragraphs', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.text('First'));
      doc.addParagraph(DocxParagraph.text('Second'));
      doc.addParagraph(DocxParagraph.text('Third'));

      final result = _roundTrip(doc);
      expect(result.paragraphs, hasLength(3));
      expect(result.paragraphs[0].plainText, 'First');
      expect(result.paragraphs[1].plainText, 'Second');
      expect(result.paragraphs[2].plainText, 'Third');
    });

    test('reads empty paragraph', () {
      final doc = DocxDocument();
      doc.addParagraph(const DocxParagraph(runs: []));

      final result = _roundTrip(doc);
      expect(result.paragraphs, hasLength(1));
      expect(result.paragraphs.first.runs, isEmpty);
    });
  });

  group('DocxReader - round-trip: headings', () {
    test('reads heading levels 1-4', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.heading('H1', level: 1));
      doc.addParagraph(DocxParagraph.heading('H2', level: 2));
      doc.addParagraph(DocxParagraph.heading('H3', level: 3));
      doc.addParagraph(DocxParagraph.heading('H4', level: 4));

      final result = _roundTrip(doc);
      expect(result.paragraphs[0].style, DocxParagraphStyle.heading1);
      expect(result.paragraphs[1].style, DocxParagraphStyle.heading2);
      expect(result.paragraphs[2].style, DocxParagraphStyle.heading3);
      expect(result.paragraphs[3].style, DocxParagraphStyle.heading4);
    });
  });

  group('DocxReader - round-trip: paragraph styles', () {
    test('reads subtitle', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.subtitle('Sub'));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.style, DocxParagraphStyle.subtitle);
    });

    test('reads caption', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.caption('Cap'));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.style, DocxParagraphStyle.caption);
    });

    test('reads quote', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.quote('Quoted'));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.style, DocxParagraphStyle.quote);
    });

    test('reads code block', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.codeBlock('print("hi")'));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.style, DocxParagraphStyle.codeBlock);
    });

    test('reads footnote', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.footnote('Note'));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.style, DocxParagraphStyle.footnote);
    });
  });

  group('DocxReader - round-trip: alignment', () {
    test('reads center alignment', () {
      final doc = DocxDocument();
      doc.addParagraph(
          DocxParagraph.text('Centered', alignment: DocxAlignment.center));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.alignment, DocxAlignment.center);
    });

    test('reads right alignment', () {
      final doc = DocxDocument();
      doc.addParagraph(
          DocxParagraph.text('Right', alignment: DocxAlignment.right));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.alignment, DocxAlignment.right);
    });

    test('reads justify alignment', () {
      final doc = DocxDocument();
      doc.addParagraph(
          DocxParagraph.text('Justified', alignment: DocxAlignment.justify));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.alignment, DocxAlignment.justify);
    });
  });

  group('DocxReader - round-trip: formatting', () {
    test('reads bold', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph(
        runs: [const DocxRun('Bold', bold: true)],
      ));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.runs.first.bold, isTrue);
    });

    test('reads italic', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph(
        runs: [const DocxRun('Italic', italic: true)],
      ));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.runs.first.italic, isTrue);
    });

    test('reads underline', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph(
        runs: [const DocxRun('Underlined', underline: true)],
      ));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.runs.first.underline, isTrue);
    });

    test('reads strikethrough', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph(
        runs: [const DocxRun('Struck', strikethrough: true)],
      ));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.runs.first.strikethrough, isTrue);
    });

    test('reads text color', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph(
        runs: [const DocxRun('Red', color: 'FF0000')],
      ));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.runs.first.color, 'FF0000');
    });

    test('reads background color', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph(
        runs: [const DocxRun('Highlighted', backgroundColor: 'FFFF00')],
      ));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.runs.first.backgroundColor, 'FFFF00');
    });

    test('reads mixed formatting in one paragraph', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph(
        runs: [
          const DocxRun('Normal '),
          const DocxRun('bold', bold: true),
          const DocxRun(' and '),
          const DocxRun('italic', italic: true),
        ],
      ));

      final result = _roundTrip(doc);
      final runs = result.paragraphs.first.runs;
      expect(runs, hasLength(4));
      expect(runs[0].text, 'Normal ');
      expect(runs[0].bold, isFalse);
      expect(runs[1].text, 'bold');
      expect(runs[1].bold, isTrue);
      expect(runs[2].text, ' and ');
      expect(runs[3].text, 'italic');
      expect(runs[3].italic, isTrue);
    });

    test('reads combined formatting on single run', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph(
        runs: [
          const DocxRun('All', bold: true, italic: true, underline: true),
        ],
      ));

      final result = _roundTrip(doc);
      final run = result.paragraphs.first.runs.first;
      expect(run.bold, isTrue);
      expect(run.italic, isTrue);
      expect(run.underline, isTrue);
    });
  });

  group('DocxReader - round-trip: lists', () {
    test('reads bullet list', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.bulletItem('Item 1'));
      doc.addParagraph(DocxParagraph.bulletItem('Item 2'));

      final result = _roundTrip(doc);
      expect(result.paragraphs[0].style, DocxParagraphStyle.listBullet);
      expect(result.paragraphs[1].style, DocxParagraphStyle.listBullet);
    });

    test('reads numbered list', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.numberedItem('First'));
      doc.addParagraph(DocxParagraph.numberedItem('Second'));

      final result = _roundTrip(doc);
      expect(result.paragraphs[0].style, DocxParagraphStyle.listNumber);
      expect(result.paragraphs[1].style, DocxParagraphStyle.listNumber);
    });

    test('reads dash list', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.dashItem('Dash 1'));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.style, DocxParagraphStyle.listDash);
    });

    test('reads alpha list', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.alphaItem('Alpha'));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.style, DocxParagraphStyle.listNumberAlpha);
    });

    test('reads roman list', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.romanItem('Roman'));

      final result = _roundTrip(doc);
      expect(
          result.paragraphs.first.style, DocxParagraphStyle.listNumberRoman);
    });

    test('reads nested list with indent levels', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.bulletItem('Level 0'));
      doc.addParagraph(DocxParagraph.bulletItem('Level 1', indentLevel: 1));
      doc.addParagraph(DocxParagraph.bulletItem('Level 2', indentLevel: 2));

      final result = _roundTrip(doc);
      expect(result.paragraphs[0].indentLevel, 0);
      expect(result.paragraphs[1].indentLevel, 1);
      expect(result.paragraphs[2].indentLevel, 2);
    });
  });

  group('DocxReader - round-trip: tables', () {
    test('reads simple table', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable.simple([
        ['A', 'B'],
        ['C', 'D'],
      ]));

      final result = _roundTrip(doc);
      expect(result.tables, hasLength(1));
      final table = result.tables.first;
      expect(table.rowCount, 2);

      expect(table.rows[0].cells[0].paragraphs.first.plainText, 'A');
      expect(table.rows[0].cells[1].paragraphs.first.plainText, 'B');
      expect(table.rows[1].cells[0].paragraphs.first.plainText, 'C');
      expect(table.rows[1].cells[1].paragraphs.first.plainText, 'D');
    });

    test('reads table borders', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable.simple(
        [
          ['A']
        ],
        borders: const DocxTableBorders.all(),
      ));

      final result = _roundTrip(doc);
      final borders = result.tables.first.borders;
      expect(borders.top, isNotNull);
      expect(borders.bottom, isNotNull);
      expect(borders.left, isNotNull);
      expect(borders.right, isNotNull);
      expect(borders.insideH, isNotNull);
      expect(borders.insideV, isNotNull);
    });

    test('reads table with colspan', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable(
        rows: [
          DocxTableRow(cells: [
            DocxTableCell.text('Span 2', colSpan: 2),
          ]),
          const DocxTableRow(cells: [
            DocxTableCell(
              paragraphs: [DocxParagraph(runs: [DocxRun('A')])],
            ),
            DocxTableCell(
              paragraphs: [DocxParagraph(runs: [DocxRun('B')])],
            ),
          ]),
        ],
      ));

      final result = _roundTrip(doc);
      expect(result.tables.first.rows[0].cells[0].colSpan, 2);
    });

    test('reads table with rowspan', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable(
        rows: [
          DocxTableRow(cells: [
            DocxTableCell.text('Span 2 rows', rowSpan: 2),
            DocxTableCell.text('B'),
          ]),
          const DocxTableRow(cells: [
            DocxTableCell.merged(),
            DocxTableCell(
              paragraphs: [DocxParagraph(runs: [DocxRun('D')])],
            ),
          ]),
        ],
      ));

      final result = _roundTrip(doc);
      expect(result.tables.first.rows[0].cells[0].rowSpan, 2);
      expect(result.tables.first.rows[1].cells[0].isMergedContinuation, isTrue);
    });

    test('reads cell background color', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable(
        rows: [
          DocxTableRow(cells: [
            DocxTableCell.text('Colored', backgroundColor: 'E0E0E0'),
          ]),
        ],
      ));

      final result = _roundTrip(doc);
      expect(result.tables.first.rows[0].cells[0].backgroundColor, 'E0E0E0');
    });

    test('reads cell vertical alignment', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable(
        rows: [
          DocxTableRow(cells: [
            const DocxTableCell(
              paragraphs: [DocxParagraph(runs: [DocxRun('Center')])],
              verticalAlignment: DocxVerticalAlignment.center,
            ),
          ]),
        ],
      ));

      final result = _roundTrip(doc);
      expect(
        result.tables.first.rows[0].cells[0].verticalAlignment,
        DocxVerticalAlignment.center,
      );
    });
  });

  group('DocxReader - round-trip: hyperlinks', () {
    test('reads external hyperlink', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph(
        runs: [
          const DocxRun('Click here', hyperlink: 'https://example.com'),
        ],
      ));

      final result = _roundTrip(doc);
      final run = result.paragraphs.first.runs.first;
      expect(run.hyperlink, 'https://example.com');
      expect(run.text, 'Click here');
    });

    test('reads internal bookmark reference', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.heading(
        'Target',
        level: 1,
        bookmarkName: 'my_bookmark',
      ));
      doc.addParagraph(DocxParagraph(
        runs: [
          const DocxRun('Go to target', bookmarkRef: 'my_bookmark'),
        ],
      ));

      final result = _roundTrip(doc);
      expect(result.paragraphs.first.bookmarkName, 'my_bookmark');
      expect(result.paragraphs[1].runs.first.bookmarkRef, 'my_bookmark');
    });
  });

  group('DocxReader - round-trip: page breaks', () {
    test('reads page break before paragraph', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.text('Before'));
      doc.addParagraph(
          DocxParagraph.text('After break', pageBreakBefore: true));

      final result = _roundTrip(doc);
      expect(result.paragraphs, hasLength(2));
      expect(result.paragraphs[0].pageBreakBefore, isFalse);
      expect(result.paragraphs[1].pageBreakBefore, isTrue);
      expect(result.paragraphs[1].plainText, 'After break');
    });
  });

  group('DocxReader - round-trip: line breaks', () {
    test('reads line break run', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph(
        runs: [
          const DocxRun('Before'),
          const DocxRun.lineBreak(),
          const DocxRun('After'),
        ],
      ));

      final result = _roundTrip(doc);
      final runs = result.paragraphs.first.runs;
      expect(runs, hasLength(3));
      expect(runs[0].text, 'Before');
      expect(runs[1].isLineBreak, isTrue);
      expect(runs[2].text, 'After');
    });
  });

  group('DocxReader - round-trip: mixed content order', () {
    test('preserves paragraph-table-paragraph order', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.text('Before table'));
      doc.addTable(DocxTable.simple([
        ['A', 'B']
      ]));
      doc.addParagraph(DocxParagraph.text('After table'));

      final result = _roundTrip(doc);
      final content = result.content;
      expect(content, hasLength(3));
      expect(content[0], isA<DocxParagraph>());
      expect((content[0] as DocxParagraph).plainText, 'Before table');
      expect(content[1], isA<DocxTable>());
      expect(content[2], isA<DocxParagraph>());
      expect((content[2] as DocxParagraph).plainText, 'After table');
    });

    test('preserves multiple tables interspersed with paragraphs', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.text('P1'));
      doc.addTable(DocxTable.simple([
        ['T1']
      ]));
      doc.addParagraph(DocxParagraph.text('P2'));
      doc.addTable(DocxTable.simple([
        ['T2']
      ]));
      doc.addParagraph(DocxParagraph.text('P3'));

      final result = _roundTrip(doc);
      expect(result.content, hasLength(5));
      expect(result.content[0], isA<DocxParagraph>());
      expect(result.content[1], isA<DocxTable>());
      expect(result.content[2], isA<DocxParagraph>());
      expect(result.content[3], isA<DocxTable>());
      expect(result.content[4], isA<DocxParagraph>());
    });
  });

  group('DocxReader - round-trip: metadata', () {
    test('reads document with no metadata', () {
      final doc = DocxDocument();
      doc.addParagraph(DocxParagraph.text('Hello'));

      final result = _roundTrip(doc);
      expect(result.title, isNull);
      expect(result.author, isNull);
    });
  });

  group('DocxReader - round-trip: cell borders', () {
    test('reads per-cell borders', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable(
        rows: [
          DocxTableRow(cells: [
            const DocxTableCell(
              paragraphs: [DocxParagraph(runs: [DocxRun('With borders')])],
              borders: DocxCellBorders.all(),
            ),
          ]),
        ],
      ));

      final result = _roundTrip(doc);
      final cellBorders = result.tables.first.rows[0].cells[0].borders;
      expect(cellBorders, isNotNull);
      expect(cellBorders!.top, isNotNull);
      expect(cellBorders.bottom, isNotNull);
      expect(cellBorders.left, isNotNull);
      expect(cellBorders.right, isNotNull);
    });
  });
}

import 'package:docs_gee/docs_gee.dart';
import 'package:docs_gee/src/xml_parts/document_xml.dart';
import 'package:test/test.dart';

void main() {
  group('DocxCellBorders', () {
    group('constructors', () {
      test('.all() creates borders on all four sides', () {
        const borders = DocxCellBorders.all();

        expect(borders.top, isNotNull);
        expect(borders.bottom, isNotNull);
        expect(borders.left, isNotNull);
        expect(borders.right, isNotNull);
        expect(borders.hasBorders, isTrue);
      });

      test('.all() uses default color and size', () {
        const borders = DocxCellBorders.all();

        expect(borders.top!.color, '000000');
        expect(borders.top!.size, 4);
        expect(borders.top!.style, DocxBorderStyle.single);
      });

      test('.none() creates no borders', () {
        const borders = DocxCellBorders.none();

        expect(borders.top, isNull);
        expect(borders.bottom, isNull);
        expect(borders.left, isNull);
        expect(borders.right, isNull);
        expect(borders.hasBorders, isFalse);
      });

      test('.bottom() creates only a bottom border', () {
        const borders = DocxCellBorders.bottom();

        expect(borders.top, isNull);
        expect(borders.bottom, isNotNull);
        expect(borders.left, isNull);
        expect(borders.right, isNull);
        expect(borders.hasBorders, isTrue);
      });

      test('.bottom() accepts custom border', () {
        const customBorder = DocxBorder(
          color: 'FF0000',
          size: 8,
          style: DocxBorderStyle.double,
        );
        const borders = DocxCellBorders.bottom(border: customBorder);

        expect(borders.bottom!.color, 'FF0000');
        expect(borders.bottom!.size, 8);
        expect(borders.bottom!.style, DocxBorderStyle.double);
      });

      test('named constructor allows selective sides', () {
        const borders = DocxCellBorders(
          top: DocxBorder(color: 'FF0000'),
          right: DocxBorder(color: '0000FF'),
        );

        expect(borders.top, isNotNull);
        expect(borders.bottom, isNull);
        expect(borders.left, isNull);
        expect(borders.right, isNotNull);
        expect(borders.hasBorders, isTrue);
      });
    });

    group('hasBorders', () {
      test('returns true when only top is set', () {
        const borders = DocxCellBorders(top: DocxBorder());
        expect(borders.hasBorders, isTrue);
      });

      test('returns true when only left is set', () {
        const borders = DocxCellBorders(left: DocxBorder());
        expect(borders.hasBorders, isTrue);
      });

      test('returns false when no borders set', () {
        const borders = DocxCellBorders();
        expect(borders.hasBorders, isFalse);
      });
    });
  });

  group('DocxTableCell with borders', () {
    test('default cell has null borders', () {
      const cell = DocxTableCell();
      expect(cell.borders, isNull);
    });

    test('.text() factory accepts borders parameter', () {
      final cell = DocxTableCell.text(
        'Hello',
        borders: const DocxCellBorders.all(),
      );

      expect(cell.borders, isNotNull);
      expect(cell.borders!.hasBorders, isTrue);
    });

    test('constructor accepts borders parameter', () {
      const cell = DocxTableCell(
        borders: DocxCellBorders(
          bottom: DocxBorder(color: 'FF0000', size: 12),
        ),
      );

      expect(cell.borders, isNotNull);
      expect(cell.borders!.bottom!.color, 'FF0000');
      expect(cell.borders!.bottom!.size, 12);
    });

    test('.merged() has null borders', () {
      const cell = DocxTableCell.merged();
      expect(cell.borders, isNull);
    });
  });

  group('CellBorders type alias', () {
    test('CellBorders is alias for DocxCellBorders', () {
      const borders = CellBorders.all();
      expect(borders, isA<DocxCellBorders>());
      expect(borders.hasBorders, isTrue);
    });
  });

  group('DOCX XML generation with cell borders', () {
    test('cell borders produce <w:tcBorders> in XML', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable(
        borders: const DocxTableBorders.none(),
        rows: [
          DocxTableRow(cells: [
            DocxTableCell.text(
              'Bordered',
              borders: const DocxCellBorders.all(),
            ),
            DocxTableCell.text('No borders'),
          ]),
        ],
      ));

      final result = DocumentXml.generate(doc);
      final xml = result.xml;

      expect(xml, contains('<w:tcBorders>'));
      // Only the first cell should have tcBorders
      final firstTcBorders = xml.indexOf('<w:tcBorders>');
      final secondTcBorders = xml.indexOf('<w:tcBorders>', firstTcBorders + 1);
      expect(secondTcBorders, -1);
    });

    test('cell borders contain correct border elements', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable(
        borders: const DocxTableBorders.none(),
        rows: [
          DocxTableRow(cells: [
            DocxTableCell.text(
              'Custom',
              borders: const DocxCellBorders(
                top: DocxBorder(
                  color: 'FF0000',
                  size: 8,
                  style: DocxBorderStyle.double,
                ),
                bottom: DocxBorder(color: '00FF00'),
              ),
            ),
          ]),
        ],
      ));

      final result = DocumentXml.generate(doc);
      final xml = result.xml;

      expect(xml, contains('<w:tcBorders>'));
      expect(
        xml,
        contains('w:top w:val="double" w:sz="8" w:space="0" w:color="FF0000"'),
      );
      expect(
        xml,
        contains(
            'w:bottom w:val="single" w:sz="4" w:space="0" w:color="00FF00"'),
      );
      // left and right are null, so they should be "nil"
      expect(xml, contains('<w:left w:val="nil"/>'));
      expect(xml, contains('<w:right w:val="nil"/>'));
    });

    test('cell without borders produces no <w:tcBorders>', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable(
        borders: const DocxTableBorders.none(),
        rows: [
          DocxTableRow(cells: [
            DocxTableCell.text('Plain'),
          ]),
        ],
      ));

      final result = DocumentXml.generate(doc);
      expect(result.xml, isNot(contains('<w:tcBorders>')));
    });

    test('cell borders coexist with table-level borders', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable(
        borders: const DocxTableBorders.all(),
        rows: [
          DocxTableRow(cells: [
            DocxTableCell.text(
              'Cell override',
              borders: const DocxCellBorders(
                bottom: DocxBorder(
                  color: 'FF0000',
                  size: 16,
                  style: DocxBorderStyle.dashed,
                ),
              ),
            ),
            DocxTableCell.text('Table borders'),
          ]),
        ],
      ));

      final result = DocumentXml.generate(doc);
      final xml = result.xml;

      // Table-level borders should exist
      expect(xml, contains('<w:tblBorders>'));
      // Cell-level borders should also exist
      expect(xml, contains('<w:tcBorders>'));
      expect(
        xml,
        contains(
            'w:bottom w:val="dashed" w:sz="16" w:space="0" w:color="FF0000"'),
      );
    });

    test('cell borders with background color and vertical alignment', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable(
        rows: [
          DocxTableRow(cells: [
            DocxTableCell.text(
              'Styled',
              borders: const DocxCellBorders.all(),
              backgroundColor: 'FFFF00',
              verticalAlignment: DocxVerticalAlignment.center,
            ),
          ]),
        ],
      ));

      final result = DocumentXml.generate(doc);
      final xml = result.xml;

      expect(xml, contains('<w:tcBorders>'));
      expect(xml, contains('<w:shd w:val="clear" w:fill="FFFF00"/>'));
      expect(xml, contains('<w:vAlign w:val="center"/>'));
    });
  });

  group('PDF generation with cell borders', () {
    test('generates PDF without errors when cell borders are set', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable(
        borders: const DocxTableBorders.none(),
        rows: [
          DocxTableRow(cells: [
            DocxTableCell.text(
              'Red border',
              borders: const DocxCellBorders.all(color: 'FF0000', size: 8),
            ),
            DocxTableCell.text('No borders'),
          ]),
          DocxTableRow(cells: [
            DocxTableCell.text('Default'),
            DocxTableCell.text(
              'Blue bottom',
              borders: const DocxCellBorders.bottom(
                border: DocxBorder(color: '0000FF'),
              ),
            ),
          ]),
        ],
      ));

      final pdfGenerator = PdfGenerator();
      final bytes = pdfGenerator.generate(doc);

      expect(bytes, isNotEmpty);
      // PDF header
      expect(String.fromCharCodes(bytes.take(5)), '%PDF-');
    });

    test('generates PDF without errors for mixed table and cell borders', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable(
        borders: const DocxTableBorders.all(),
        rows: [
          DocxTableRow(cells: [
            DocxTableCell.text('Table borders'),
            DocxTableCell.text(
              'Cell override',
              borders: const DocxCellBorders(
                top: DocxBorder(color: 'FF0000', size: 16),
              ),
            ),
          ]),
        ],
      ));

      final pdfGenerator = PdfGenerator();
      final bytes = pdfGenerator.generate(doc);

      expect(bytes, isNotEmpty);
    });
  });

  group('DOCX generation with cell borders', () {
    test('generates DOCX without errors when cell borders are set', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable(
        rows: [
          DocxTableRow(cells: [
            DocxTableCell.text(
              'Bordered',
              borders: const DocxCellBorders.all(color: 'FF0000'),
            ),
            DocxTableCell.text('Normal'),
          ]),
        ],
      ));

      final docxGenerator = DocxGenerator();
      final bytes = docxGenerator.generate(doc);

      expect(bytes, isNotEmpty);
      // ZIP/DOCX magic bytes (PK)
      expect(bytes[0], 0x50); // P
      expect(bytes[1], 0x4B); // K
    });
  });

  group('border styles in cell borders', () {
    for (final style in DocxBorderStyle.values) {
      test('${style.name} style produces correct XML value', () {
        final doc = DocxDocument();
        doc.addTable(DocxTable(
          borders: const DocxTableBorders.none(),
          rows: [
            DocxTableRow(cells: [
              DocxTableCell.text(
                'Test',
                borders: DocxCellBorders(
                  top: DocxBorder(style: style),
                ),
              ),
            ]),
          ],
        ));

        final result = DocumentXml.generate(doc);
        expect(result.xml, contains('w:val="${style.value}"'));
      });
    }
  });
}

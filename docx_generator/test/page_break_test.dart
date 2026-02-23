import 'package:docs_gee/docs_gee.dart';
import 'package:docs_gee/src/xml_parts/document_xml.dart';
import 'package:test/test.dart';

void main() {
  group('Page break in paragraphs', () {
    test('pageBreakBefore inserts <w:br w:type="page"/> before paragraph', () {
      final doc = DocxDocument();
      doc.addParagraph(Paragraph.text('First page content'));
      doc.addParagraph(
        Paragraph.text('Second page content', pageBreakBefore: true),
      );

      final result = DocumentXml.generate(doc);
      final xml = result.xml;

      expect(xml, contains('<w:br w:type="page"/>'));
    });

    test('pageBreakBefore creates a separate paragraph with the break', () {
      final doc = DocxDocument();
      doc.addParagraph(
        Paragraph.text('After break', pageBreakBefore: true),
      );

      final result = DocumentXml.generate(doc);
      final xml = result.xml;

      // The break should be in its own <w:p> with a <w:r> containing <w:br>
      expect(xml, contains('<w:p>\n      <w:r>\n        <w:br w:type="page"/>'));
    });

    test('pageBreakBefore does not use <w:pageBreakBefore/> property', () {
      final doc = DocxDocument();
      doc.addParagraph(
        Paragraph.text('Content', pageBreakBefore: true),
      );

      final result = DocumentXml.generate(doc);
      final xml = result.xml;

      expect(xml, isNot(contains('<w:pageBreakBefore/>')));
    });

    test('paragraph without pageBreakBefore has no break element', () {
      final doc = DocxDocument();
      doc.addParagraph(Paragraph.text('Normal paragraph'));

      final result = DocumentXml.generate(doc);
      final xml = result.xml;

      expect(xml, isNot(contains('<w:br w:type="page"/>')));
      expect(xml, isNot(contains('<w:pageBreakBefore/>')));
    });

    test('heading with pageBreakBefore inserts break before heading', () {
      final doc = DocxDocument();
      doc.addParagraph(
        Paragraph.heading('Chapter 2', level: 1, pageBreakBefore: true),
      );

      final result = DocumentXml.generate(doc);
      final xml = result.xml;

      expect(xml, contains('<w:br w:type="page"/>'));
      // Heading style should still be present in the actual paragraph
      expect(xml, contains('w:val="Heading1"'));
    });

    test('multiple page breaks produce multiple break elements', () {
      final doc = DocxDocument();
      doc.addParagraph(Paragraph.text('Page 1'));
      doc.addParagraph(
        Paragraph.text('Page 2', pageBreakBefore: true),
      );
      doc.addParagraph(
        Paragraph.text('Page 3', pageBreakBefore: true),
      );

      final result = DocumentXml.generate(doc);
      final xml = result.xml;

      final matches = '<w:br w:type="page"/>'
          .allMatches(xml)
          .length;
      expect(matches, 2);
    });
  });

  group('Page break in table cell paragraphs', () {
    test('pageBreakBefore in cell paragraph inserts break element', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable(
        rows: [
          DocxTableRow(cells: [
            DocxTableCell(paragraphs: [
              Paragraph.text('Before break'),
              Paragraph.text('After break', pageBreakBefore: true),
            ]),
          ]),
        ],
      ));

      final result = DocumentXml.generate(doc);
      final xml = result.xml;

      expect(xml, contains('<w:br w:type="page"/>'));
    });

    test('cell paragraph pageBreakBefore does not use property element', () {
      final doc = DocxDocument();
      doc.addTable(DocxTable(
        rows: [
          DocxTableRow(cells: [
            DocxTableCell(paragraphs: [
              Paragraph.text('Content', pageBreakBefore: true),
            ]),
          ]),
        ],
      ));

      final result = DocumentXml.generate(doc);
      final xml = result.xml;

      expect(xml, isNot(contains('<w:pageBreakBefore/>')));
    });
  });

  group('Page break after TOC', () {
    test('TOC uses <w:br w:type="page"/> not <w:pageBreakBefore/>', () {
      final doc = DocxDocument(includeTableOfContents: true);
      doc.addParagraph(Paragraph.heading('Chapter', level: 1));

      final result = DocumentXml.generate(doc);
      final xml = result.xml;

      expect(xml, contains('<w:br w:type="page"/>'));
      expect(xml, isNot(contains('<w:pageBreakBefore/>')));
    });
  });

  group('DOCX generation with page breaks', () {
    test('generates valid DOCX with page breaks', () {
      final doc = DocxDocument();
      doc.addParagraph(Paragraph.text('Page 1'));
      doc.addParagraph(
        Paragraph.heading('Page 2', level: 1, pageBreakBefore: true),
      );
      doc.addParagraph(Paragraph.text('Content on page 2'));
      doc.addParagraph(
        Paragraph.text('Page 3', pageBreakBefore: true),
      );

      final docxGenerator = DocxGenerator();
      final bytes = docxGenerator.generate(doc);

      expect(bytes, isNotEmpty);
      // ZIP/DOCX magic bytes (PK)
      expect(bytes[0], 0x50);
      expect(bytes[1], 0x4B);
    });
  });

  group('PDF generation with page breaks', () {
    test('generates valid PDF with page breaks', () {
      final doc = DocxDocument();
      doc.addParagraph(Paragraph.text('Page 1'));
      doc.addParagraph(
        Paragraph.text('Page 2', pageBreakBefore: true),
      );

      final pdfGenerator = PdfGenerator();
      final bytes = pdfGenerator.generate(doc);

      expect(bytes, isNotEmpty);
      expect(String.fromCharCodes(bytes.take(5)), '%PDF-');
    });
  });
}

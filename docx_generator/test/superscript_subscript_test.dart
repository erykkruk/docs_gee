import 'package:docs_gee/docs_gee.dart';
import 'package:docs_gee/src/xml_parts/document_xml.dart';
import 'package:test/test.dart';

void main() {
  group('DocxScript on DocxRun', () {
    test('baseline is the default script', () {
      const run = DocxRun('text');

      expect(run.script, DocxScript.baseline);
    });

    test('script counts as formatting for superscript', () {
      const run = DocxRun('2', script: DocxScript.superscript);

      expect(run.hasFormatting, isTrue);
    });

    test('script counts as formatting for subscript', () {
      const run = DocxRun('2', script: DocxScript.subscript);

      expect(run.hasFormatting, isTrue);
    });

    test('baseline script does not count as formatting on its own', () {
      const run = DocxRun('text');

      expect(run.hasFormatting, isFalse);
    });

    test('copyWith updates the script', () {
      const run = DocxRun('x');

      final updated = run.copyWith(script: DocxScript.superscript);

      expect(updated.script, DocxScript.superscript);
      expect(updated.text, 'x');
    });

    test('copyWith preserves the script when omitted', () {
      const run = DocxRun('x', script: DocxScript.subscript);

      final updated = run.copyWith(bold: true);

      expect(updated.script, DocxScript.subscript);
      expect(updated.bold, isTrue);
    });
  });

  group('DOCX superscript/subscript rendering', () {
    test('superscript run emits <w:vertAlign w:val="superscript"/>', () {
      final doc = DocxDocument();
      doc.addParagraph(const DocxParagraph(runs: [
        DocxRun('E=mc'),
        DocxRun('2', script: DocxScript.superscript),
      ]));

      final xml = DocumentXml.generate(doc).xml;

      expect(xml, contains('<w:vertAlign w:val="superscript"/>'));
    });

    test('subscript run emits <w:vertAlign w:val="subscript"/>', () {
      final doc = DocxDocument();
      doc.addParagraph(const DocxParagraph(runs: [
        DocxRun('H'),
        DocxRun('2', script: DocxScript.subscript),
        DocxRun('O'),
      ]));

      final xml = DocumentXml.generate(doc).xml;

      expect(xml, contains('<w:vertAlign w:val="subscript"/>'));
    });

    test('baseline run does not emit a vertAlign element', () {
      final doc = DocxDocument();
      doc.addParagraph(const DocxParagraph(runs: [DocxRun('plain')]));

      final xml = DocumentXml.generate(doc).xml;

      expect(xml, isNot(contains('<w:vertAlign')));
    });

    test('superscript works inside table cells', () {
      final doc = DocxDocument();
      doc.addTable(const DocxTable(
        rows: [
          DocxTableRow(cells: [
            DocxTableCell(paragraphs: [
              DocxParagraph(runs: [
                DocxRun('n'),
                DocxRun('th', script: DocxScript.superscript),
              ]),
            ]),
          ]),
        ],
      ));

      final xml = DocumentXml.generate(doc).xml;

      expect(xml, contains('<w:vertAlign w:val="superscript"/>'));
    });
  });

  group('DOCX round-trip of scripts', () {
    test('superscript survives write then read', () {
      final doc = DocxDocument();
      doc.addParagraph(const DocxParagraph(runs: [
        DocxRun('x'),
        DocxRun('2', script: DocxScript.superscript),
      ]));

      final bytes = DocxGenerator().generate(doc);
      final restored = const DocxReader().read(bytes);

      final runs = restored.paragraphs.single.runs;
      expect(runs.last.text, '2');
      expect(runs.last.script, DocxScript.superscript);
    });

    test('subscript survives write then read', () {
      final doc = DocxDocument();
      doc.addParagraph(const DocxParagraph(runs: [
        DocxRun('CO'),
        DocxRun('2', script: DocxScript.subscript),
      ]));

      final bytes = DocxGenerator().generate(doc);
      final restored = const DocxReader().read(bytes);

      final runs = restored.paragraphs.single.runs;
      expect(runs.last.script, DocxScript.subscript);
    });
  });

  group('PDF superscript/subscript rendering', () {
    test('superscript applies a positive text rise (Ts operator)', () {
      final doc = DocxDocument();
      doc.addParagraph(const DocxParagraph(runs: [
        DocxRun('E=mc'),
        DocxRun('2', script: DocxScript.superscript),
      ]));

      final bytes = PdfGenerator().generate(doc);
      final content = String.fromCharCodes(bytes);

      expect(content, contains(' Ts'));
    });

    test('generates a valid PDF with subscript content', () {
      final doc = DocxDocument();
      doc.addParagraph(const DocxParagraph(runs: [
        DocxRun('H'),
        DocxRun('2', script: DocxScript.subscript),
        DocxRun('O'),
      ]));

      final bytes = PdfGenerator().generate(doc);

      expect(bytes, isNotEmpty);
      expect(String.fromCharCodes(bytes.take(5)), '%PDF-');
    });

    test('baseline-only document contains no text rise', () {
      final doc = DocxDocument();
      doc.addParagraph(const DocxParagraph(runs: [DocxRun('plain text')]));

      final bytes = PdfGenerator().generate(doc);
      final content = String.fromCharCodes(bytes);

      expect(content, isNot(contains(' Ts')));
    });
  });
}

import 'package:docs_gee/docs_gee.dart';
import 'package:test/test.dart';

import 'support/docx_archive.dart';

void main() {
  group('DocxHeaderFooter - model', () {
    test('.text() wraps the string in a single paragraph', () {
      final header = DocxHeaderFooter.text('Quarterly report');

      expect(header.paragraphs, hasLength(1));
      expect(header.paragraphs.first.runs.first.text, 'Quarterly report');
      expect(header.isEmpty, isFalse);
    });

    test('.text() applies the requested alignment', () {
      final header = DocxHeaderFooter.text(
        'Centered',
        alignment: DocxAlignment.center,
      );

      expect(header.paragraphs.first.alignment, DocxAlignment.center);
    });

    test('.pageNumber() emits a single PAGE field by default', () {
      final footer = DocxHeaderFooter.pageNumber();
      final runs = footer.paragraphs.first.runs;

      expect(runs, hasLength(1));
      expect(runs.single.field, DocxField.page);
    });

    test('.pageNumber() with showTotal emits both fields and a separator', () {
      final footer = DocxHeaderFooter.pageNumber(
        prefix: 'Page ',
        showTotal: true,
      );
      final runs = footer.paragraphs.first.runs;

      expect(runs.map((r) => r.field).toList(), [
        null,
        DocxField.page,
        null,
        DocxField.pageCount,
      ]);
      expect(runs.first.text, 'Page ');
      expect(runs[2].text, ' of ');
    });

    test('.pageNumber() honours a custom separator and suffix', () {
      final footer = DocxHeaderFooter.pageNumber(
        showTotal: true,
        separator: ' / ',
        suffix: '.',
      );
      final texts = footer.paragraphs.first.runs.map((r) => r.text).toList();

      expect(texts, contains(' / '));
      expect(texts.last, '.');
    });

    test('an empty header reports isEmpty', () {
      const header = DocxHeaderFooter(paragraphs: []);

      expect(header.isEmpty, isTrue);
    });
  });

  group('DocxRun - field constructors', () {
    test('pageNumber carries the PAGE instruction', () {
      const run = DocxRun.pageNumber();

      expect(run.isField, isTrue);
      expect(run.field!.instruction, 'PAGE');
    });

    test('pageCount carries the NUMPAGES instruction', () {
      const run = DocxRun.pageCount();

      expect(run.isField, isTrue);
      expect(run.field!.instruction, 'NUMPAGES');
    });

    test('a plain run is not a field', () {
      expect(const DocxRun('text').isField, isFalse);
    });

    test('field runs accept their own formatting', () {
      const run = DocxRun.pageNumber(bold: true, fontSize: 9, color: '888888');

      expect(run.bold, isTrue);
      expect(run.fontSize, 9);
      expect(run.color, '888888');
    });
  });

  group('DOCX generation - header and footer parts', () {
    test('writes header1.xml when a header is set', () {
      final doc = DocxDocument(header: DocxHeaderFooter.text('Top'))
        ..addParagraph(DocxParagraph.text('Body'));

      final bytes = DocxGenerator().generate(doc);

      expect(partNames(bytes), contains('word/header1.xml'));
      expect(partText(bytes, 'word/header1.xml'), contains('<w:hdr'));
      expect(partText(bytes, 'word/header1.xml'), contains('Top'));
    });

    test('writes footer1.xml when a footer is set', () {
      final doc = DocxDocument(footer: DocxHeaderFooter.text('Bottom'))
        ..addParagraph(DocxParagraph.text('Body'));

      final bytes = DocxGenerator().generate(doc);

      expect(partNames(bytes), contains('word/footer1.xml'));
      expect(partText(bytes, 'word/footer1.xml'), contains('<w:ftr'));
    });

    test('writes neither part for a document without them', () {
      final doc = DocxDocument()..addParagraph(DocxParagraph.text('Body'));

      final names = partNames(DocxGenerator().generate(doc));

      expect(names, isNot(contains('word/header1.xml')));
      expect(names, isNot(contains('word/footer1.xml')));
    });

    test('an empty header adds no part and no relationship', () {
      final doc = DocxDocument(header: const DocxHeaderFooter(paragraphs: []))
        ..addParagraph(DocxParagraph.text('Body'));

      final bytes = DocxGenerator().generate(doc);

      expect(partNames(bytes), isNot(contains('word/header1.xml')));
      expect(
        partText(bytes, 'word/_rels/document.xml.rels'),
        isNot(contains('relationships/header')),
      );
    });

    test('declares the content type overrides for both parts', () {
      final doc = DocxDocument(
        header: DocxHeaderFooter.text('Top'),
        footer: DocxHeaderFooter.pageNumber(),
      )..addParagraph(DocxParagraph.text('Body'));

      final types =
          partText(DocxGenerator().generate(doc), '[Content_Types].xml')!;

      expect(types, contains('/word/header1.xml'));
      expect(types, contains('/word/footer1.xml'));
      expect(types, contains('wordprocessingml.header+xml'));
      expect(types, contains('wordprocessingml.footer+xml'));
    });

    test('sectPr references resolve against the relationship file', () {
      final doc = DocxDocument(
        header: DocxHeaderFooter.text('Top'),
        footer: DocxHeaderFooter.text('Bottom'),
      )..addParagraph(DocxParagraph.text('Body'));

      final bytes = DocxGenerator().generate(doc);
      final document = partText(bytes, 'word/document.xml')!;
      final rels = partText(bytes, 'word/_rels/document.xml.rels')!;

      final headerRef = RegExp(r'<w:headerReference[^>]*r:id="(rId\d+)"')
          .firstMatch(document);
      final footerRef = RegExp(r'<w:footerReference[^>]*r:id="(rId\d+)"')
          .firstMatch(document);

      expect(headerRef, isNotNull);
      expect(footerRef, isNotNull);
      expect(rels, contains('Id="${headerRef!.group(1)}"'));
      expect(rels, contains('Id="${footerRef!.group(1)}"'));
    });

    test('header and footer references come first inside sectPr', () {
      // The schema fixes this order; Word rejects the file otherwise.
      final doc = DocxDocument(header: DocxHeaderFooter.text('Top'))
        ..addParagraph(DocxParagraph.text('Body'));

      final document =
          partText(DocxGenerator().generate(doc), 'word/document.xml')!;

      expect(
        document.indexOf('<w:headerReference'),
        lessThan(document.indexOf('<w:pgSz')),
      );
    });

    test('a page-number footer emits a PAGE field', () {
      final doc = DocxDocument(footer: DocxHeaderFooter.pageNumber())
        ..addParagraph(DocxParagraph.text('Body'));

      final footer =
          partText(DocxGenerator().generate(doc), 'word/footer1.xml')!;

      expect(footer, contains('w:instr=" PAGE "'));
    });

    test('a "page X of Y" footer emits both field instructions', () {
      final doc = DocxDocument(
        footer: DocxHeaderFooter.pageNumber(prefix: 'Page ', showTotal: true),
      )..addParagraph(DocxParagraph.text('Body'));

      final footer =
          partText(DocxGenerator().generate(doc), 'word/footer1.xml')!;

      expect(footer, contains('w:instr=" PAGE "'));
      expect(footer, contains('w:instr=" NUMPAGES "'));
      expect(footer, contains('Page '));
    });

    test('a hyperlink inside a header degrades to plain text', () {
      // Header parts resolve r:id against their own relationship file, which
      // this library does not emit; a dangling id would break the document.
      final doc = DocxDocument(
        header: const DocxHeaderFooter(
          paragraphs: [
            DocxParagraph(
              runs: [DocxRun('docs', hyperlink: 'https://example.com')],
            ),
          ],
        ),
      )..addParagraph(DocxParagraph.text('Body'));

      final header =
          partText(DocxGenerator().generate(doc), 'word/header1.xml')!;

      expect(header, isNot(contains('r:id=')));
      expect(header, contains('docs'));
    });

    test('body hyperlinks still get their relationship', () {
      final doc = DocxDocument(header: DocxHeaderFooter.text('Top'))
        ..addParagraph(
          const DocxParagraph(
            runs: [DocxRun('docs', hyperlink: 'https://example.com')],
          ),
        );

      final bytes = DocxGenerator().generate(doc);

      expect(partText(bytes, 'word/document.xml'), contains('r:id='));
      expect(
        partText(bytes, 'word/_rels/document.xml.rels'),
        contains('https://example.com'),
      );
    });

    test('a header with several paragraphs writes them all', () {
      final doc = DocxDocument(
        header: const DocxHeaderFooter(
          paragraphs: [
            DocxParagraph(runs: [DocxRun('First line')]),
            DocxParagraph(runs: [DocxRun('Second line')]),
          ],
        ),
      )..addParagraph(DocxParagraph.text('Body'));

      final header =
          partText(DocxGenerator().generate(doc), 'word/header1.xml')!;

      expect(header, contains('First line'));
      expect(header, contains('Second line'));
    });
  });
}

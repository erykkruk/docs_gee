import 'dart:convert';
import 'dart:io';

import 'package:docs_gee/docs_gee.dart';

/// Demonstrates the 1.5.0 features: embedded images, headers and footers with
/// page numbers, right-to-left text, per-run font sizes and cell padding.
///
/// Run with: `dart run example/images_headers_example.dart`
void main() {
  final doc = DocxDocument(
    title: 'docs_gee 1.5.0 feature tour',
    author: 'docs_gee',
    header: DocxHeaderFooter.text(
      'Quarterly report',
      alignment: DocxAlignment.center,
    ),
    // Renders "Page 1 of 3"; Word recalculates the fields when it opens.
    footer: DocxHeaderFooter.pageNumber(prefix: 'Page ', showTotal: true),
  );

  doc.addParagraph(DocxParagraph.heading('Images', level: 1));
  doc.addParagraph(DocxParagraph.text(
    'The intrinsic size comes from the PNG header, so no dimensions are '
    'needed. Pass width or height to scale while keeping the aspect ratio.',
  ));
  doc.addImage(DocxImage(
    bytes: base64Decode(_onePixelPng),
    width: 120,
    altText: 'A placeholder square',
    alignment: DocxAlignment.center,
  ));

  doc.addParagraph(DocxParagraph.heading('Font size per run', level: 1));
  doc.addParagraph(const DocxParagraph(runs: [
    DocxRun('This lead-in is 18pt, ', fontSize: 18),
    DocxRun('and the rest runs at the document default.'),
  ]));

  doc.addParagraph(DocxParagraph.heading('Right-to-left text', level: 1));
  doc.addParagraph(const DocxParagraph(
    runs: [DocxRun('مرحبا بالعالم', rtl: true)],
    rtl: true,
  ));
  doc.addParagraph(const DocxParagraph(
    runs: [DocxRun('שלום עולם', rtl: true)],
    rtl: true,
  ));

  doc.addParagraph(DocxParagraph.heading('Table cell padding', level: 1));
  doc.addTable(DocxTable.fromHeaders(
    headers: ['Product', 'Price'],
    rows: [
      ['Standard plan', '19.00'],
      ['Pro plan', '49.00'],
    ],
    cellPadding: const DocxCellPadding.points(
      top: 6,
      bottom: 6,
      left: 10,
      right: 10,
    ),
  ));

  final docx = DocxGenerator().generate(doc);
  File('example_1_5_0.docx').writeAsBytesSync(docx);
  stdout.writeln('Wrote example_1_5_0.docx (${docx.length} bytes)');

  // Headers, footers and images are DOCX-only; the PDF keeps the text,
  // per-run sizes and cell padding.
  final pdf = PdfGenerator().generate(doc);
  File('example_1_5_0.pdf').writeAsBytesSync(pdf);
  stdout.writeln('Wrote example_1_5_0.pdf (${pdf.length} bytes)');
}

/// A valid 1x1 PNG, used so the example needs no asset files on disk.
const String _onePixelPng =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4z8AAAAM'
    'BAQAY3Y2wAAAAAElFTkSuQmCC';

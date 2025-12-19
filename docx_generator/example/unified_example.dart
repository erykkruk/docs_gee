/// Example demonstrating the unified DocumentGenerator interface.
///
/// This example shows how to use the same document model
/// to export to both DOCX and PDF formats.
///
/// Run with: dart run example/unified_example.dart

// ignore_for_file: avoid_print

import 'dart:io';

import 'package:docs_gee/docs_gee.dart';

void main() {
  // Create document using generic aliases
  final doc = Document(
    title: 'Unified Generator Demo',
    author: 'docs_gee library',
  );

  // Add content using generic Paragraph alias
  doc.addParagraph(Paragraph.heading(
    'Unified Document Generator',
    level: 1,
    alignment: Alignment.center,
  ));

  doc.addParagraph(Paragraph.text(''));

  doc.addParagraph(Paragraph.text(
    'This document was created using the unified DocumentGenerator interface. '
    'The same document model can be exported to both DOCX and PDF formats.',
  ));

  doc.addParagraph(Paragraph.heading('Features', level: 2));

  doc.addParagraph(Paragraph.bulletItem('Single document model'));
  doc.addParagraph(Paragraph.bulletItem('Multiple export formats'));
  doc.addParagraph(Paragraph.bulletItem('Interchangeable generators'));
  doc.addParagraph(Paragraph.bulletItem('Generic type aliases'));

  doc.addParagraph(Paragraph.heading('Formatted Text', level: 2));

  doc.addParagraph(Paragraph(
    runs: [
      const TextRun('You can use '),
      const TextRun('bold', bold: true),
      const TextRun(', '),
      const TextRun('italic', italic: true),
      const TextRun(', and '),
      const TextRun('combined styles', bold: true, italic: true),
      const TextRun('.'),
    ],
  ));

  // Export to both formats using the same interface
  exportDocument(doc, DocxGenerator(), 'unified_demo.docx');
  exportDocument(doc, PdfGenerator(), 'unified_demo.pdf');

  print('Done! Generated both DOCX and PDF from the same document.');
}

/// Exports a document using any DocumentGenerator implementation.
void exportDocument(Document doc, DocumentGenerator generator, String filename) {
  final bytes = generator.generate(doc);
  File(filename).writeAsBytesSync(bytes);
  print('Generated: $filename (${bytes.length} bytes)');
}

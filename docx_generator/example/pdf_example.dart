/// Example demonstrating PDF generation using docs_gee library.
///
/// This example shows how to create a PDF document with various features:
/// - Headings (H1, H2, H3)
/// - Text formatting (bold, italic)
/// - Text alignment (left, center, right, justify)
/// - Bullet lists
/// - Numbered lists
/// - Page breaks
///
/// Run with: dart run example/pdf_example.dart

// ignore_for_file: avoid_print

import 'dart:io';
import 'dart:typed_data';

import 'package:docs_gee/docs_gee.dart';

void main() {
  final bytes = generateSamplePdfDocument();

  // Save to file
  final outputPath = 'pdf_generator_demo.pdf';
  File(outputPath).writeAsBytesSync(bytes);

  print('PDF document generated successfully!');
  print('Output: $outputPath');
  print('Size: ${bytes.length} bytes');
}

/// Generates a sample PDF document demonstrating all features.
///
/// Returns the document as bytes, which can be:
/// - Saved to a file (desktop/mobile)
/// - Downloaded via browser (web)
/// - Uploaded to a server
/// - Attached to an email
Uint8List generateSamplePdfDocument() {
  // Create document with metadata
  // Note: Using the same DocxDocument model for both DOCX and PDF generation
  final doc = DocxDocument(
    title: 'PDF Generator Demo',
    author: 'docs_gee library',
  );

  // ============================================================
  // TITLE PAGE
  // ============================================================

  doc.addParagraph(DocxParagraph.heading(
    'PDF Generator Library',
    level: 1,
    alignment: DocxAlignment.center,
  ));

  doc.addParagraph(DocxParagraph.text(
    'A demonstration of PDF generation capabilities',
    alignment: DocxAlignment.center,
  ));

  doc.addParagraph(DocxParagraph.text(''));

  doc.addParagraph(DocxParagraph.text(
    'This document was generated programmatically using the docs_gee '
    'Dart library. It demonstrates various formatting options available '
    'for PDF generation without any external dependencies.',
    alignment: DocxAlignment.justify,
  ));

  // ============================================================
  // TEXT FORMATTING SECTION
  // ============================================================

  doc.addParagraph(DocxParagraph.heading(
    'Text Formatting',
    level: 1,
    pageBreakBefore: true,
  ));

  doc.addParagraph(DocxParagraph.text(
    'The library supports various text formatting options. '
    'You can apply formatting to individual text runs within a paragraph.',
  ));

  doc.addParagraph(DocxParagraph.heading('Basic Formatting', level: 2));

  // Mixed formatting example
  doc.addParagraph(DocxParagraph(
    runs: [
      const DocxRun('This paragraph contains '),
      const DocxRun('bold text', bold: true),
      const DocxRun(', '),
      const DocxRun('italic text', italic: true),
      const DocxRun(', and '),
      const DocxRun('bold italic text', bold: true, italic: true),
      const DocxRun('.'),
    ],
  ));

  doc.addParagraph(DocxParagraph.heading('Combined Formatting', level: 2));

  doc.addParagraph(DocxParagraph(
    runs: [
      const DocxRun('PDF uses the standard '),
      const DocxRun('Base 14 fonts', bold: true),
      const DocxRun(' (Helvetica, Times-Roman, Courier) which support '),
      const DocxRun('bold', bold: true),
      const DocxRun(' and '),
      const DocxRun('italic', italic: true),
      const DocxRun(' variants.'),
    ],
  ));

  // ============================================================
  // TEXT ALIGNMENT SECTION
  // ============================================================

  doc.addParagraph(DocxParagraph.heading(
    'Text Alignment',
    level: 1,
    pageBreakBefore: true,
  ));

  doc.addParagraph(DocxParagraph.text(
    'LEFT ALIGNED: This paragraph is aligned to the left margin. '
    'This is the default alignment for most text content.',
    alignment: DocxAlignment.left,
  ));

  doc.addParagraph(DocxParagraph.text(''));

  doc.addParagraph(DocxParagraph.text(
    'CENTER ALIGNED: This paragraph is centered between the margins. '
    'Useful for titles and headings.',
    alignment: DocxAlignment.center,
  ));

  doc.addParagraph(DocxParagraph.text(''));

  doc.addParagraph(DocxParagraph.text(
    'RIGHT ALIGNED: This paragraph is aligned to the right margin. '
    'Often used for dates or signatures.',
    alignment: DocxAlignment.right,
  ));

  // ============================================================
  // HEADINGS SECTION
  // ============================================================

  doc.addParagraph(DocxParagraph.heading(
    'Heading Levels',
    level: 1,
    pageBreakBefore: true,
  ));

  doc.addParagraph(DocxParagraph.text(
    'The library supports three heading levels for document structure:',
  ));

  doc.addParagraph(DocxParagraph.heading('Heading Level 1', level: 1));
  doc.addParagraph(DocxParagraph.text(
    'Primary headings for main sections.',
  ));

  doc.addParagraph(DocxParagraph.heading('Heading Level 2', level: 2));
  doc.addParagraph(DocxParagraph.text(
    'Secondary headings for subsections.',
  ));

  doc.addParagraph(DocxParagraph.heading('Heading Level 3', level: 3));
  doc.addParagraph(DocxParagraph.text(
    'Tertiary headings for sub-subsections.',
  ));

  // ============================================================
  // LISTS SECTION
  // ============================================================

  doc.addParagraph(DocxParagraph.heading(
    'Lists',
    level: 1,
    pageBreakBefore: true,
  ));

  doc.addParagraph(DocxParagraph.heading('Bullet List', level: 2));

  doc.addParagraph(DocxParagraph.text('Features of this library:'));

  doc.addParagraph(DocxParagraph.bulletItem('Pure Dart implementation'));
  doc.addParagraph(DocxParagraph.bulletItem('Zero external dependencies for PDF'));
  doc.addParagraph(DocxParagraph.bulletItem('Cross-platform support'));
  doc.addParagraph(DocxParagraph.bulletItem('Same API for DOCX and PDF'));
  doc.addParagraph(DocxParagraph.bulletItem('Lightweight footprint'));

  doc.addParagraph(DocxParagraph.heading('Numbered List', level: 2));

  doc.addParagraph(DocxParagraph.text('Steps to generate a PDF:'));

  doc.addParagraph(DocxParagraph.numberedItem('Create a DocxDocument'));
  doc.addParagraph(DocxParagraph.numberedItem('Add paragraphs with content'));
  doc.addParagraph(DocxParagraph.numberedItem('Create a PdfGenerator'));
  doc.addParagraph(DocxParagraph.numberedItem('Call generate() to get bytes'));
  doc.addParagraph(DocxParagraph.numberedItem('Save or transmit the bytes'));

  // ============================================================
  // CONCLUSION
  // ============================================================

  doc.addParagraph(DocxParagraph.heading(
    'Conclusion',
    level: 1,
    pageBreakBefore: true,
  ));

  doc.addParagraph(DocxParagraph.text(
    'This document demonstrates the core capabilities of the docs_gee '
    'PDF generator. The generated PDF file is compatible with all standard '
    'PDF readers including Adobe Acrobat, Preview, Chrome, and others.',
    alignment: DocxAlignment.justify,
  ));

  doc.addParagraph(DocxParagraph.text(''));

  doc.addParagraph(DocxParagraph(
    runs: [
      const DocxRun('Key advantage: '),
      const DocxRun('same document model', bold: true),
      const DocxRun(' can be exported to both DOCX and PDF formats!'),
    ],
    alignment: DocxAlignment.center,
  ));

  // Generate PDF with custom font settings
  final generator = PdfGenerator(
    fontName: 'Helvetica',
    fontSize: 12,
  );

  return generator.generate(doc);
}

/// Example demonstrating docx_generator library usage.
///
/// This example shows how to create a DOCX document with various features:
/// - Headings (H1, H2, H3)
/// - Text formatting (bold, italic, underline, strikethrough)
/// - Text alignment (left, center, right, justify)
/// - Bullet lists
/// - Numbered lists
/// - Page breaks
///
/// Run with: dart run example/docx_generator_example.dart

// ignore_for_file: avoid_print

import 'dart:io';
import 'dart:typed_data';

import 'package:docs_gee/docs_gee.dart';

void main() {
  final bytes = generateSampleDocument();

  // Save to file
  final outputPath = 'docx_generator_demo.docx';
  File(outputPath).writeAsBytesSync(bytes);

  print('Document generated successfully!');
  print('Output: $outputPath');
  print('Size: ${bytes.length} bytes');
}

/// Generates a sample DOCX document demonstrating all features.
///
/// Returns the document as bytes, which can be:
/// - Saved to a file (desktop/mobile)
/// - Downloaded via browser (web)
/// - Uploaded to a server
/// - Attached to an email
Uint8List generateSampleDocument() {
  // Create document with metadata
  final doc = DocxDocument(
    title: 'DOCX Generator Demo',
    author: 'docx_generator library',
  );

  // ============================================================
  // TITLE PAGE
  // ============================================================

  doc.addParagraph(DocxParagraph.heading(
    'DOCX Generator Library',
    level: 1,
    alignment: DocxAlignment.center,
  ));

  doc.addParagraph(DocxParagraph.text(
    'A demonstration of document generation capabilities',
    alignment: DocxAlignment.center,
  ));

  doc.addParagraph(DocxParagraph.text(''));

  doc.addParagraph(DocxParagraph.text(
    'This document was generated programmatically using the docx_generator '
    'Dart library. It demonstrates various formatting options available.',
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
      const DocxRun(', '),
      const DocxRun('underlined text', underline: true),
      const DocxRun(', and '),
      const DocxRun('strikethrough text', strikethrough: true),
      const DocxRun('.'),
    ],
  ));

  doc.addParagraph(DocxParagraph.heading('Combined Formatting', level: 2));

  doc.addParagraph(DocxParagraph(
    runs: [
      const DocxRun('You can also '),
      const DocxRun('combine multiple styles', bold: true, italic: true),
      const DocxRun(' in a single run, like '),
      const DocxRun(
        'bold + italic + underline',
        bold: true,
        italic: true,
        underline: true,
      ),
      const DocxRun('.'),
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

  doc.addParagraph(DocxParagraph.text(''));

  doc.addParagraph(DocxParagraph.text(
    'JUSTIFIED: This paragraph is justified, meaning the text is spread '
    'evenly between both margins. This creates clean edges on both sides '
    'and is commonly used in books, newspapers, and formal documents. '
    'The spacing between words is adjusted to achieve this effect.',
    alignment: DocxAlignment.justify,
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
  doc.addParagraph(DocxParagraph.bulletItem('No native dependencies'));
  doc.addParagraph(DocxParagraph.bulletItem('Cross-platform support'));
  doc.addParagraph(DocxParagraph.bulletItem('Easy to use API'));
  doc.addParagraph(DocxParagraph.bulletItem('Lightweight footprint'));

  doc.addParagraph(DocxParagraph.heading('Numbered List', level: 2));

  doc.addParagraph(DocxParagraph.text('Steps to generate a document:'));

  doc.addParagraph(DocxParagraph.numberedItem('Create a DocxDocument'));
  doc.addParagraph(DocxParagraph.numberedItem('Add paragraphs with content'));
  doc.addParagraph(DocxParagraph.numberedItem('Create a DocxGenerator'));
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
    'This document demonstrates the core capabilities of the docx_generator '
    'library. The generated DOCX file is compatible with Microsoft Word, '
    'Google Docs, LibreOffice Writer, and other word processors that '
    'support the OOXML format.',
    alignment: DocxAlignment.justify,
  ));

  doc.addParagraph(DocxParagraph.text(''));

  doc.addParagraph(DocxParagraph.text(
    'For more information, visit the project repository.',
    alignment: DocxAlignment.center,
  ));

  // Generate DOCX with custom font settings
  final generator = DocxGenerator(
    fontName: 'Times New Roman',
    fontSize: 24, // 12pt
  );

  return generator.generate(doc);
}

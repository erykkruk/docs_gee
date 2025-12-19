/// Comprehensive example demonstrating ALL features of docs_gee.
///
/// This example shows every available feature including:
/// - Text formatting (bold, italic, underline, strikethrough, colors)
/// - All heading levels (H1-H4)
/// - Semantic styles (subtitle, caption, quote)
/// - All list types (bullet, dash, numbered, alpha, roman)
/// - Nested lists
/// - Page breaks
///
/// Run with: dart run example/all_features_example.dart

// ignore_for_file: avoid_print

import 'dart:io';

import 'package:docs_gee/docs_gee.dart';

void main() {
  final doc = Document(
    title: 'All Features Demo',
    author: 'docs_gee library',
  );

  // ============================================
  // HEADINGS
  // ============================================
  doc.addParagraph(Paragraph.heading('1. Headings', level: 1));
  doc.addParagraph(Paragraph.heading('Heading Level 1', level: 1));
  doc.addParagraph(Paragraph.heading('Heading Level 2', level: 2));
  doc.addParagraph(Paragraph.heading('Heading Level 3', level: 3));
  doc.addParagraph(Paragraph.heading('Heading Level 4', level: 4));
  doc.addParagraph(Paragraph.text(''));

  // ============================================
  // SEMANTIC STYLES
  // ============================================
  doc.addParagraph(Paragraph.heading('2. Semantic Styles', level: 1));
  doc.addParagraph(
      Paragraph.subtitle('This is a subtitle - great for document subtitles'));
  doc.addParagraph(
      Paragraph.caption('This is a caption - perfect for image captions'));
  doc.addParagraph(Paragraph.quote(
    'This is a blockquote - ideal for citations and quotes. '
    'It appears indented and in italic style.',
  ));
  doc.addParagraph(Paragraph.text(''));

  // ============================================
  // TEXT FORMATTING
  // ============================================
  doc.addParagraph(Paragraph.heading('3. Text Formatting', level: 1));

  doc.addParagraph(Paragraph(
    runs: [
      const TextRun('Regular text, '),
      const TextRun('bold text, ', bold: true),
      const TextRun('italic text, ', italic: true),
      const TextRun('underlined text, ', underline: true),
      const TextRun('strikethrough text.', strikethrough: true),
    ],
  ));

  doc.addParagraph(Paragraph(
    runs: [
      const TextRun('Combined: ', bold: true),
      const TextRun('bold + italic', bold: true, italic: true),
      const TextRun(', '),
      const TextRun('bold + underline', bold: true, underline: true),
      const TextRun(', '),
      const TextRun('all styles',
          bold: true, italic: true, underline: true, strikethrough: true),
    ],
  ));
  doc.addParagraph(Paragraph.text(''));

  // ============================================
  // TEXT COLORS
  // ============================================
  doc.addParagraph(Paragraph.heading('4. Text Colors', level: 1));

  doc.addParagraph(Paragraph(
    runs: [
      const TextRun('Red text ', color: 'FF0000'),
      const TextRun('Green text ', color: '00FF00'),
      const TextRun('Blue text ', color: '0000FF'),
      const TextRun('Orange text ', color: 'FFA500'),
      const TextRun('Purple text', color: '800080'),
    ],
  ));

  doc.addParagraph(Paragraph(
    runs: [
      const TextRun('Yellow highlight ', backgroundColor: 'FFFF00'),
      const TextRun('Cyan highlight ', backgroundColor: '00FFFF'),
      const TextRun('Magenta highlight', backgroundColor: 'FF00FF'),
    ],
  ));
  doc.addParagraph(Paragraph.text(''));

  // ============================================
  // TEXT ALIGNMENT
  // ============================================
  doc.addParagraph(Paragraph.heading('5. Text Alignment', level: 1));
  doc.addParagraph(
      Paragraph.text('Left aligned text (default)', alignment: Alignment.left));
  doc.addParagraph(
      Paragraph.text('Center aligned text', alignment: Alignment.center));
  doc.addParagraph(
      Paragraph.text('Right aligned text', alignment: Alignment.right));
  doc.addParagraph(Paragraph.text(
    'Justified text fills the entire line width by adjusting spacing between words. '
    'This creates a clean, professional look for longer paragraphs.',
    alignment: Alignment.justify,
  ));
  doc.addParagraph(Paragraph.text(''));

  // ============================================
  // BULLET LISTS
  // ============================================
  doc.addParagraph(Paragraph.heading('6. Unordered Lists', level: 1));

  doc.addParagraph(Paragraph.heading('Bullet List:', level: 3));
  doc.addParagraph(Paragraph.bulletItem('First bullet item'));
  doc.addParagraph(Paragraph.bulletItem('Second bullet item'));
  doc.addParagraph(Paragraph.bulletItem('Third bullet item'));

  doc.addParagraph(Paragraph.heading('Dash List:', level: 3));
  doc.addParagraph(Paragraph.dashItem('First dash item'));
  doc.addParagraph(Paragraph.dashItem('Second dash item'));
  doc.addParagraph(Paragraph.dashItem('Third dash item'));
  doc.addParagraph(Paragraph.text(''));

  // ============================================
  // NUMBERED LISTS
  // ============================================
  doc.addParagraph(Paragraph.heading('7. Ordered Lists', level: 1));

  doc.addParagraph(Paragraph.heading('Numeric List (1, 2, 3):', level: 3));
  doc.addParagraph(Paragraph.numberedItem('First numbered item'));
  doc.addParagraph(Paragraph.numberedItem('Second numbered item'));
  doc.addParagraph(Paragraph.numberedItem('Third numbered item'));

  doc.addParagraph(Paragraph.heading('Alphabetic List (a, b, c):', level: 3));
  doc.addParagraph(Paragraph.alphaItem('First alpha item'));
  doc.addParagraph(Paragraph.alphaItem('Second alpha item'));
  doc.addParagraph(Paragraph.alphaItem('Third alpha item'));

  doc.addParagraph(
      Paragraph.heading('Roman Numeral List (I, II, III):', level: 3));
  doc.addParagraph(Paragraph.romanItem('First roman item'));
  doc.addParagraph(Paragraph.romanItem('Second roman item'));
  doc.addParagraph(Paragraph.romanItem('Third roman item'));
  doc.addParagraph(Paragraph.text(''));

  // ============================================
  // NESTED LISTS
  // ============================================
  doc.addParagraph(Paragraph.heading('8. Nested Lists', level: 1));

  doc.addParagraph(Paragraph.bulletItem('Top level item 1'));
  doc.addParagraph(Paragraph.bulletItem('Nested item 1.1', indentLevel: 1));
  doc.addParagraph(Paragraph.bulletItem('Nested item 1.2', indentLevel: 1));
  doc.addParagraph(Paragraph.bulletItem('Deep nested 1.2.1', indentLevel: 2));
  doc.addParagraph(Paragraph.bulletItem('Deep nested 1.2.2', indentLevel: 2));
  doc.addParagraph(Paragraph.bulletItem('Top level item 2'));
  doc.addParagraph(Paragraph.bulletItem('Nested item 2.1', indentLevel: 1));

  doc.addParagraph(Paragraph.text(''));

  doc.addParagraph(Paragraph.heading('Mixed Nested Lists:', level: 3));
  doc.addParagraph(Paragraph.numberedItem('First main item'));
  doc.addParagraph(Paragraph.alphaItem('Sub-item a', indentLevel: 1));
  doc.addParagraph(Paragraph.alphaItem('Sub-item b', indentLevel: 1));
  doc.addParagraph(Paragraph.romanItem('Detail i', indentLevel: 2));
  doc.addParagraph(Paragraph.romanItem('Detail ii', indentLevel: 2));
  doc.addParagraph(Paragraph.numberedItem('Second main item'));
  doc.addParagraph(Paragraph.dashItem('Note', indentLevel: 1));
  doc.addParagraph(Paragraph.text(''));

  // ============================================
  // PAGE BREAK
  // ============================================
  doc.addParagraph(Paragraph.heading(
    '9. Page Break Demo',
    level: 1,
    pageBreakBefore: true,
  ));
  doc.addParagraph(Paragraph.text(
    'This section starts on a new page! '
    'Page breaks are useful for separating chapters or sections.',
  ));

  // ============================================
  // SUMMARY
  // ============================================
  doc.addParagraph(
      Paragraph.heading('Summary', level: 1, pageBreakBefore: true));
  doc.addParagraph(Paragraph.text(
    'This document demonstrates all the features available in docs_gee library:',
  ));
  doc.addParagraph(Paragraph.bulletItem('4 heading levels (H1-H4)'));
  doc.addParagraph(
      Paragraph.bulletItem('Semantic styles: subtitle, caption, quote'));
  doc.addParagraph(Paragraph.bulletItem(
      'Text formatting: bold, italic, underline, strikethrough'));
  doc.addParagraph(Paragraph.bulletItem('Text colors and background colors'));
  doc.addParagraph(
      Paragraph.bulletItem('Text alignment: left, center, right, justify'));
  doc.addParagraph(Paragraph.bulletItem(
      '5 list types: bullet, dash, numbered, alpha, roman'));
  doc.addParagraph(Paragraph.bulletItem('Nested lists up to 9 levels deep'));
  doc.addParagraph(Paragraph.bulletItem('Page breaks'));
  doc.addParagraph(
      Paragraph.bulletItem('Document metadata: title, author, dates'));

  // Export to both formats
  final docxGenerator = DocxGenerator();
  final pdfGenerator = PdfGenerator();

  final docxBytes = docxGenerator.generate(doc);
  final pdfBytes = pdfGenerator.generate(doc);

  File('all_features.docx').writeAsBytesSync(docxBytes);
  File('all_features.pdf').writeAsBytesSync(pdfBytes);

  print('Generated all_features.docx (${docxBytes.length} bytes)');
  print('Generated all_features.pdf (${pdfBytes.length} bytes)');
  print('');
  print('Open the files to see all features in action!');
}

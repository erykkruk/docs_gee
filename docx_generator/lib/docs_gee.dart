/// Pure Dart library for generating DOCX and PDF files.
///
/// This library allows you to create Word documents (.docx) and PDF files
/// without external dependencies beyond the `archive` package for DOCX ZIP creation.
/// PDF generation is completely dependency-free.
///
/// Example using generic aliases:
/// ```dart
/// final doc = Document();
/// doc.addParagraph(Paragraph.heading('Title', level: 1));
/// doc.addParagraph(Paragraph.text('Hello World'));
///
/// // Export to any format
/// DocumentGenerator generator = PdfGenerator(); // or DocxGenerator()
/// final bytes = generator.generate(doc);
/// ```
library docs_gee;

export 'src/document_generator.dart';
export 'src/docx_generator.dart';
export 'src/pdf_generator.dart';
export 'src/models/models.dart';

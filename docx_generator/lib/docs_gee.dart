/// Pure Dart library for generating DOCX and PDF files.
///
/// This library allows you to create Word documents (.docx) and PDF files
/// without external dependencies beyond the `archive` package for DOCX ZIP creation.
/// PDF generation is completely dependency-free.
///
/// **Platform Support:**
/// - Android, iOS, macOS, Windows, Linux: Full support
/// - Web: Full support (use browser APIs for file saving)
///
/// Example using generic aliases:
/// ```dart
/// final doc = Document();
/// doc.addParagraph(Paragraph.heading('Title', level: 1));
/// doc.addParagraph(Paragraph.text('Hello World'));
///
/// // Export to any format (works on ALL platforms including Web)
/// DocumentGenerator generator = PdfGenerator(); // or DocxGenerator()
/// final bytes = generator.generate(doc);
///
/// // Saving files (non-Web platforms only):
/// // import 'package:docs_gee/src/file_saver_io.dart';
/// // await bytes.saveToFile('output.pdf');
/// ```
library docs_gee;

export 'src/document_generator.dart';
export 'src/docx_generator.dart';
export 'src/pdf_generator.dart';
export 'src/models/models.dart';

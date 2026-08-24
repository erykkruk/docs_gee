import 'docx_enums.dart';
import 'docx_paragraph.dart';
import 'docx_run.dart';

/// Content repeated at the top or bottom of every page.
///
/// A header or footer holds ordinary paragraphs, so any run formatting works
/// inside it. Combine it with the field runs on [DocxRun] to get live page
/// numbers:
///
/// ```dart
/// final doc = DocxDocument(
///   header: DocxHeaderFooter.text('Quarterly report', alignment: DocxAlignment.center),
///   footer: DocxHeaderFooter.pageNumber(prefix: 'Page '),
/// );
/// ```
///
/// Word recalculates the page fields when the document opens, so the numbers
/// are correct without the generator ever paginating the content itself.
class DocxHeaderFooter {
  /// Creates a header or footer from explicit paragraphs.
  const DocxHeaderFooter({required this.paragraphs});

  /// Creates a single-paragraph header or footer with plain text.
  factory DocxHeaderFooter.text(
    String text, {
    DocxAlignment alignment = DocxAlignment.left,
  }) {
    return DocxHeaderFooter(
      paragraphs: [
        DocxParagraph(runs: [DocxRun(text)], alignment: alignment),
      ],
    );
  }

  /// Creates a footer (or header) showing the current page number.
  ///
  /// With [showTotal] the text reads `Page 3 of 12` using the [separator]
  /// between the two fields; without it, just the page number.
  factory DocxHeaderFooter.pageNumber({
    String prefix = '',
    String suffix = '',
    bool showTotal = false,
    String separator = ' of ',
    DocxAlignment alignment = DocxAlignment.center,
  }) {
    return DocxHeaderFooter(
      paragraphs: [
        DocxParagraph(
          runs: [
            if (prefix.isNotEmpty) DocxRun(prefix),
            const DocxRun.pageNumber(),
            if (showTotal) ...[
              DocxRun(separator),
              const DocxRun.pageCount(),
            ],
            if (suffix.isNotEmpty) DocxRun(suffix),
          ],
          alignment: alignment,
        ),
      ],
    );
  }

  /// Paragraphs rendered inside the header or footer, in order.
  final List<DocxParagraph> paragraphs;

  /// Whether this header or footer would render nothing.
  bool get isEmpty => paragraphs.isEmpty;
}

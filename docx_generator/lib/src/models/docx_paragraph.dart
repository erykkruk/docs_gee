import 'docx_enums.dart';
import 'docx_run.dart';

/// Represents a paragraph in a DOCX document.
class DocxParagraph {
  const DocxParagraph({
    required this.runs,
    this.style = DocxParagraphStyle.normal,
    this.alignment = DocxAlignment.left,
    this.pageBreakBefore = false,
    this.indentLevel = 0,
  });

  /// Creates a simple paragraph with plain text.
  factory DocxParagraph.text(
    String text, {
    DocxParagraphStyle style = DocxParagraphStyle.normal,
    DocxAlignment alignment = DocxAlignment.left,
    bool pageBreakBefore = false,
  }) {
    return DocxParagraph(
      runs: [DocxRun(text)],
      style: style,
      alignment: alignment,
      pageBreakBefore: pageBreakBefore,
    );
  }

  /// Creates a heading paragraph.
  factory DocxParagraph.heading(
    String text, {
    required int level,
    DocxAlignment alignment = DocxAlignment.left,
    bool pageBreakBefore = false,
  }) {
    final style = switch (level) {
      1 => DocxParagraphStyle.heading1,
      2 => DocxParagraphStyle.heading2,
      3 => DocxParagraphStyle.heading3,
      4 => DocxParagraphStyle.heading4,
      _ => DocxParagraphStyle.heading1,
    };
    return DocxParagraph(
      runs: [DocxRun(text)],
      style: style,
      alignment: alignment,
      pageBreakBefore: pageBreakBefore,
    );
  }

  /// Creates a subtitle paragraph.
  factory DocxParagraph.subtitle(
    String text, {
    DocxAlignment alignment = DocxAlignment.left,
  }) {
    return DocxParagraph(
      runs: [DocxRun(text)],
      style: DocxParagraphStyle.subtitle,
      alignment: alignment,
    );
  }

  /// Creates a caption paragraph.
  factory DocxParagraph.caption(
    String text, {
    DocxAlignment alignment = DocxAlignment.left,
  }) {
    return DocxParagraph(
      runs: [DocxRun(text)],
      style: DocxParagraphStyle.caption,
      alignment: alignment,
    );
  }

  /// Creates a quote/blockquote paragraph.
  factory DocxParagraph.quote(
    String text, {
    DocxAlignment alignment = DocxAlignment.left,
  }) {
    return DocxParagraph(
      runs: [DocxRun(text)],
      style: DocxParagraphStyle.quote,
      alignment: alignment,
    );
  }

  /// Creates a code block paragraph (monospace font).
  factory DocxParagraph.codeBlock(
    String text, {
    DocxAlignment alignment = DocxAlignment.left,
  }) {
    return DocxParagraph(
      runs: [DocxRun(text)],
      style: DocxParagraphStyle.codeBlock,
      alignment: alignment,
    );
  }

  /// Creates a footnote text paragraph (smaller font).
  factory DocxParagraph.footnote(
    String text, {
    DocxAlignment alignment = DocxAlignment.left,
  }) {
    return DocxParagraph(
      runs: [DocxRun(text)],
      style: DocxParagraphStyle.footnote,
      alignment: alignment,
    );
  }

  /// Creates a bullet list item (•).
  factory DocxParagraph.bulletItem(
    String text, {
    DocxAlignment alignment = DocxAlignment.left,
    int indentLevel = 0,
  }) {
    return DocxParagraph(
      runs: [DocxRun(text)],
      style: DocxParagraphStyle.listBullet,
      alignment: alignment,
      indentLevel: indentLevel,
    );
  }

  /// Creates a dash list item (-).
  factory DocxParagraph.dashItem(
    String text, {
    DocxAlignment alignment = DocxAlignment.left,
    int indentLevel = 0,
  }) {
    return DocxParagraph(
      runs: [DocxRun(text)],
      style: DocxParagraphStyle.listDash,
      alignment: alignment,
      indentLevel: indentLevel,
    );
  }

  /// Creates a numbered list item (1, 2, 3...).
  factory DocxParagraph.numberedItem(
    String text, {
    DocxAlignment alignment = DocxAlignment.left,
    int indentLevel = 0,
  }) {
    return DocxParagraph(
      runs: [DocxRun(text)],
      style: DocxParagraphStyle.listNumber,
      alignment: alignment,
      indentLevel: indentLevel,
    );
  }

  /// Creates an alphabetic list item (a, b, c...).
  factory DocxParagraph.alphaItem(
    String text, {
    DocxAlignment alignment = DocxAlignment.left,
    int indentLevel = 0,
  }) {
    return DocxParagraph(
      runs: [DocxRun(text)],
      style: DocxParagraphStyle.listNumberAlpha,
      alignment: alignment,
      indentLevel: indentLevel,
    );
  }

  /// Creates a roman numeral list item (I, II, III...).
  factory DocxParagraph.romanItem(
    String text, {
    DocxAlignment alignment = DocxAlignment.left,
    int indentLevel = 0,
  }) {
    return DocxParagraph(
      runs: [DocxRun(text)],
      style: DocxParagraphStyle.listNumberRoman,
      alignment: alignment,
      indentLevel: indentLevel,
    );
  }

  /// The text runs in this paragraph.
  final List<DocxRun> runs;

  /// The paragraph style.
  final DocxParagraphStyle style;

  /// Text alignment.
  final DocxAlignment alignment;

  /// Whether to insert a page break before this paragraph.
  final bool pageBreakBefore;

  /// Indent level for nested lists (0 = top level, 1 = first nested, etc.).
  /// Maximum supported level is 8.
  final int indentLevel;

  /// Returns the plain text content of this paragraph.
  String get plainText => runs.map((r) => r.text).join();
}

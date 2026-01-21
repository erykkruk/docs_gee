/// Represents a run of text with formatting.
///
/// A "run" in DOCX terminology is a contiguous piece of text
/// that shares the same formatting properties.
class DocxRun {
  const DocxRun(
    this.text, {
    this.bold = false,
    this.italic = false,
    this.underline = false,
    this.strikethrough = false,
    this.color,
    this.backgroundColor,
    this.hyperlink,
    this.bookmarkRef,
    this.isLineBreak = false,
  });

  /// Creates a line break run (soft return within a paragraph).
  /// This is equivalent to Shift+Enter in Word.
  const DocxRun.lineBreak()
      : text = '',
        bold = false,
        italic = false,
        underline = false,
        strikethrough = false,
        color = null,
        backgroundColor = null,
        hyperlink = null,
        bookmarkRef = null,
        isLineBreak = true;

  /// The text content.
  final String text;

  /// Whether the text is bold.
  final bool bold;

  /// Whether the text is italic.
  final bool italic;

  /// Whether the text is underlined.
  final bool underline;

  /// Whether the text has strikethrough.
  final bool strikethrough;

  /// Text color in hex format (e.g., "FF0000" for red).
  /// Without the # prefix.
  final String? color;

  /// Background/highlight color in hex format (e.g., "FFFF00" for yellow).
  /// Without the # prefix.
  final String? backgroundColor;

  /// External hyperlink URL (e.g., "https://example.com").
  /// When set, this run will be rendered as a clickable link.
  final String? hyperlink;

  /// Reference to an internal bookmark name.
  /// When set, this run will link to the bookmark within the document.
  final String? bookmarkRef;

  /// Whether this run represents a line break (soft return).
  /// When true, this generates a <w:br/> element instead of text.
  final bool isLineBreak;

  /// Returns true if this run is a link (external or internal).
  bool get isLink => hyperlink != null || bookmarkRef != null;

  /// Returns true if any formatting is applied.
  bool get hasFormatting =>
      bold ||
      italic ||
      underline ||
      strikethrough ||
      color != null ||
      backgroundColor != null;

  /// Creates a copy with modified properties.
  DocxRun copyWith({
    String? text,
    bool? bold,
    bool? italic,
    bool? underline,
    bool? strikethrough,
    String? color,
    String? backgroundColor,
    String? hyperlink,
    String? bookmarkRef,
    bool? isLineBreak,
  }) {
    return DocxRun(
      text ?? this.text,
      bold: bold ?? this.bold,
      italic: italic ?? this.italic,
      underline: underline ?? this.underline,
      strikethrough: strikethrough ?? this.strikethrough,
      color: color ?? this.color,
      backgroundColor: backgroundColor ?? this.backgroundColor,
      hyperlink: hyperlink ?? this.hyperlink,
      bookmarkRef: bookmarkRef ?? this.bookmarkRef,
      isLineBreak: isLineBreak ?? this.isLineBreak,
    );
  }
}

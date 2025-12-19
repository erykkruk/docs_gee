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
  });

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
  }) {
    return DocxRun(
      text ?? this.text,
      bold: bold ?? this.bold,
      italic: italic ?? this.italic,
      underline: underline ?? this.underline,
      strikethrough: strikethrough ?? this.strikethrough,
      color: color ?? this.color,
      backgroundColor: backgroundColor ?? this.backgroundColor,
    );
  }
}

import 'docx_enums.dart';

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
    this.script = DocxScript.baseline,
    this.isLineBreak = false,
    this.fontSize,
    this.rtl = false,
    this.field,
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
        script = DocxScript.baseline,
        isLineBreak = true,
        fontSize = null,
        rtl = false,
        field = null;

  /// Creates a run holding the current page number.
  ///
  /// Rendered as a `PAGE` field, which Word recalculates on open, so the
  /// number is right without the generator paginating anything. Intended for
  /// [DocxHeaderFooter] content; body text may use it too.
  const DocxRun.pageNumber({
    this.bold = false,
    this.italic = false,
    this.color,
    this.fontSize,
  })  : text = '',
        underline = false,
        strikethrough = false,
        backgroundColor = null,
        hyperlink = null,
        bookmarkRef = null,
        script = DocxScript.baseline,
        isLineBreak = false,
        rtl = false,
        field = DocxField.page;

  /// Creates a run holding the total page count.
  ///
  /// Rendered as a `NUMPAGES` field. Pair it with [DocxRun.pageNumber] for
  /// the usual `Page 3 of 12` footer.
  const DocxRun.pageCount({
    this.bold = false,
    this.italic = false,
    this.color,
    this.fontSize,
  })  : text = '',
        underline = false,
        strikethrough = false,
        backgroundColor = null,
        hyperlink = null,
        bookmarkRef = null,
        script = DocxScript.baseline,
        isLineBreak = false,
        rtl = false,
        field = DocxField.pageCount;

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

  /// Vertical script position (baseline, superscript or subscript).
  ///
  /// Superscript and subscript text is rendered smaller and shifted off the
  /// baseline. Example:
  /// ```dart
  /// DocxParagraph(runs: [
  ///   DocxRun('H'),
  ///   DocxRun('2', script: DocxScript.subscript),
  ///   DocxRun('O'),
  /// ]);
  /// ```
  final DocxScript script;

  /// Whether this run represents a line break (soft return).
  /// When true, this generates a `<w:br/>` element instead of text.
  final bool isLineBreak;

  /// Font size for this run in points, overriding the document default.
  ///
  /// `14` renders as 14pt. Null keeps the size configured on the generator
  /// ([DocxGenerator.fontSize] / [PdfGenerator.fontSize]), which is how every
  /// run behaved before per-run sizing existed.
  final int? fontSize;

  /// Whether this run reads right-to-left (Arabic, Hebrew, Persian).
  ///
  /// Emits `<w:rtl/>` so Word applies bidirectional layout to the run. Set
  /// [DocxParagraph.rtl] as well to flip the paragraph direction itself.
  ///
  /// DOCX only: the PDF generator has no bidirectional text shaping and
  /// renders such runs left-to-right.
  final bool rtl;

  /// Field this run renders instead of literal [text], if any.
  ///
  /// Set by the [DocxRun.pageNumber] and [DocxRun.pageCount] constructors.
  final DocxField? field;

  /// Whether this run renders a field rather than literal text.
  bool get isField => field != null;

  /// Returns true if this run is a link (external or internal).
  bool get isLink => hyperlink != null || bookmarkRef != null;

  /// Returns true if any formatting is applied.
  bool get hasFormatting =>
      bold ||
      italic ||
      underline ||
      strikethrough ||
      color != null ||
      backgroundColor != null ||
      script != DocxScript.baseline ||
      fontSize != null ||
      rtl;

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
    DocxScript? script,
    bool? isLineBreak,
    int? fontSize,
    bool? rtl,
    DocxField? field,
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
      script: script ?? this.script,
      isLineBreak: isLineBreak ?? this.isLineBreak,
      fontSize: fontSize ?? this.fontSize,
      rtl: rtl ?? this.rtl,
      field: field ?? this.field,
    );
  }
}

/// Text alignment options for paragraphs.
enum DocxAlignment {
  left('left'),
  center('center'),
  right('right'),
  justify('both');

  const DocxAlignment(this.value);
  final String value;
}

/// Paragraph style types.
enum DocxParagraphStyle {
  normal('Normal'),
  heading1('Heading1'),
  heading2('Heading2'),
  heading3('Heading3'),
  heading4('Heading4'),
  subtitle('Subtitle'),
  caption('Caption'),
  quote('Quote'),
  codeBlock('CodeBlock'),
  footnote('Footnote'),
  listBullet('ListBullet'),
  listDash('ListDash'),
  listNumber('ListNumber'),
  listNumberAlpha('ListNumberAlpha'),
  listNumberRoman('ListNumberRoman');

  const DocxParagraphStyle(this.styleId);
  final String styleId;

  /// Returns true if this style is a list style.
  bool get isList => switch (this) {
        listBullet ||
        listDash ||
        listNumber ||
        listNumberAlpha ||
        listNumberRoman =>
          true,
        _ => false,
      };

  /// Returns true if this style is an ordered list.
  bool get isOrderedList => switch (this) {
        listNumber || listNumberAlpha || listNumberRoman => true,
        _ => false,
      };

  /// Returns true if this style is an unordered list.
  bool get isUnorderedList => switch (this) {
        listBullet || listDash => true,
        _ => false,
      };
}

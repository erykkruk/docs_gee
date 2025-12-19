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
  listBullet('ListBullet'),
  listNumber('ListNumber');

  const DocxParagraphStyle(this.styleId);
  final String styleId;
}

import 'docx_paragraph.dart';

/// Represents a complete DOCX document.
class DocxDocument {
  DocxDocument({
    List<DocxParagraph>? paragraphs,
    this.title,
    this.author,
  }) : paragraphs = paragraphs ?? [];

  /// The paragraphs in this document.
  final List<DocxParagraph> paragraphs;

  /// Optional document title (metadata).
  final String? title;

  /// Optional document author (metadata).
  final String? author;

  /// Adds a paragraph to the document.
  void addParagraph(DocxParagraph paragraph) {
    paragraphs.add(paragraph);
  }

  /// Adds multiple paragraphs to the document.
  void addParagraphs(List<DocxParagraph> paragraphs) {
    this.paragraphs.addAll(paragraphs);
  }
}

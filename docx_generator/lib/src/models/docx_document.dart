import 'docx_paragraph.dart';
import 'docx_table.dart';

/// Represents a complete DOCX document.
class DocxDocument {
  DocxDocument({
    List<DocxParagraph>? paragraphs,
    this.title,
    this.author,
    this.includeTableOfContents = false,
    this.tocTitle = 'Table of Contents',
    this.tocMaxLevel = 3,
  }) : _content = paragraphs?.cast<Object>() ?? [];

  /// Internal list holding both paragraphs and tables.
  final List<Object> _content;

  /// Optional document title (metadata).
  final String? title;

  /// Optional document author (metadata).
  final String? author;

  /// Whether to include a Table of Contents at the beginning.
  final bool includeTableOfContents;

  /// Title for the Table of Contents section.
  final String tocTitle;

  /// Maximum heading level to include in TOC (1-4).
  /// Default is 3 (includes Heading1, Heading2, Heading3).
  final int tocMaxLevel;

  /// Returns all paragraphs in the document (for backward compatibility).
  List<DocxParagraph> get paragraphs =>
      _content.whereType<DocxParagraph>().toList();

  /// Returns all tables in the document.
  List<DocxTable> get tables => _content.whereType<DocxTable>().toList();

  /// Returns all content items in order (paragraphs and tables).
  List<Object> get content => List.unmodifiable(_content);

  /// Adds a paragraph to the document.
  void addParagraph(DocxParagraph paragraph) {
    _content.add(paragraph);
  }

  /// Adds multiple paragraphs to the document.
  void addParagraphs(List<DocxParagraph> paragraphs) {
    _content.addAll(paragraphs);
  }

  /// Adds a table to the document.
  void addTable(DocxTable table) {
    _content.add(table);
  }
}

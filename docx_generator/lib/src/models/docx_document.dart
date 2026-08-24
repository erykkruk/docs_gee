import 'docx_header_footer.dart';
import 'docx_image.dart';
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
    this.header,
    this.footer,
  }) : _content = paragraphs?.cast<Object>() ?? [];

  /// Internal list holding paragraphs, tables and images in insertion order.
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

  /// Content repeated at the top of every page, if any.
  ///
  /// DOCX only: the PDF generator ignores headers and footers.
  final DocxHeaderFooter? header;

  /// Content repeated at the bottom of every page, if any.
  ///
  /// Use [DocxHeaderFooter.pageNumber] for a page-number footer.
  ///
  /// DOCX only: the PDF generator ignores headers and footers.
  final DocxHeaderFooter? footer;

  /// Whether the document defines a header or a footer with content.
  bool get hasHeaderOrFooter =>
      (header != null && !header!.isEmpty) ||
      (footer != null && !footer!.isEmpty);

  /// Returns all paragraphs in the document (for backward compatibility).
  List<DocxParagraph> get paragraphs =>
      _content.whereType<DocxParagraph>().toList();

  /// Returns all tables in the document.
  List<DocxTable> get tables => _content.whereType<DocxTable>().toList();

  /// Returns all images in the document.
  List<DocxImage> get images => _content.whereType<DocxImage>().toList();

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

  /// Adds an image to the document.
  ///
  /// The image becomes its own block, rendered in the order it was added
  /// relative to paragraphs and tables.
  ///
  /// DOCX only: the PDF generator skips images.
  void addImage(DocxImage image) {
    _content.add(image);
  }
}

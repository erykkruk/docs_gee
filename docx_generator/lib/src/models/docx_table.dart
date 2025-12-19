import 'docx_enums.dart';
import 'docx_paragraph.dart';

/// Border style options for table borders.
enum DocxBorderStyle {
  single('single'),
  double('double'),
  dashed('dashed'),
  dotted('dotted');

  const DocxBorderStyle(this.value);
  final String value;
}

/// Individual border definition.
class DocxBorder {
  const DocxBorder({
    this.color = '000000',
    this.size = 4,
    this.style = DocxBorderStyle.single,
  });

  /// Border color in hex format (e.g., "000000" for black).
  final String color;

  /// Border width in eighths of a point (4 = 0.5pt, 8 = 1pt).
  final int size;

  /// Border style.
  final DocxBorderStyle style;
}

/// Border configuration for tables.
class DocxTableBorders {
  const DocxTableBorders({
    this.top,
    this.bottom,
    this.left,
    this.right,
    this.insideH,
    this.insideV,
  });

  /// All borders with default style (single line, black).
  const DocxTableBorders.all({
    String color = '000000',
    int size = 4,
  })  : top = const DocxBorder(),
        bottom = const DocxBorder(),
        left = const DocxBorder(),
        right = const DocxBorder(),
        insideH = const DocxBorder(),
        insideV = const DocxBorder();

  /// No borders.
  const DocxTableBorders.none()
      : top = null,
        bottom = null,
        left = null,
        right = null,
        insideH = null,
        insideV = null;

  /// Outside borders only (no inside grid lines).
  const DocxTableBorders.outside({
    String color = '000000',
    int size = 4,
  })  : top = const DocxBorder(),
        bottom = const DocxBorder(),
        left = const DocxBorder(),
        right = const DocxBorder(),
        insideH = null,
        insideV = null;

  final DocxBorder? top;
  final DocxBorder? bottom;
  final DocxBorder? left;
  final DocxBorder? right;

  /// Horizontal inside borders (between rows).
  final DocxBorder? insideH;

  /// Vertical inside borders (between columns).
  final DocxBorder? insideV;

  /// Returns true if any border is defined.
  bool get hasBorders =>
      top != null ||
      bottom != null ||
      left != null ||
      right != null ||
      insideH != null ||
      insideV != null;
}

/// Represents a cell in a table row.
class DocxTableCell {
  const DocxTableCell({
    this.paragraphs = const [],
    this.alignment = DocxAlignment.left,
    this.backgroundColor,
  });

  /// Creates a simple cell with plain text.
  factory DocxTableCell.text(
    String text, {
    DocxAlignment alignment = DocxAlignment.left,
    String? backgroundColor,
  }) {
    return DocxTableCell(
      paragraphs: [DocxParagraph.text(text, alignment: alignment)],
      alignment: alignment,
      backgroundColor: backgroundColor,
    );
  }

  /// The paragraphs (content) in this cell.
  final List<DocxParagraph> paragraphs;

  /// Text alignment within the cell.
  final DocxAlignment alignment;

  /// Background color in hex format (e.g., "FFFF00" for yellow).
  final String? backgroundColor;
}

/// Represents a row in a table.
class DocxTableRow {
  const DocxTableRow({
    required this.cells,
  });

  /// The cells in this row.
  final List<DocxTableCell> cells;
}

/// Represents a table in a DOCX document.
class DocxTable {
  const DocxTable({
    required this.rows,
    this.borders = const DocxTableBorders.all(),
  });

  /// The rows in this table.
  final List<DocxTableRow> rows;

  /// Table border configuration.
  final DocxTableBorders borders;

  /// Returns the number of columns (based on first row).
  int get columnCount => rows.isEmpty ? 0 : rows.first.cells.length;

  /// Returns the number of rows.
  int get rowCount => rows.length;
}

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

/// Border configuration for individual table cells.
///
/// When set on a cell, these borders override the table-level borders
/// for that specific cell.
class DocxCellBorders {
  const DocxCellBorders({
    this.top,
    this.bottom,
    this.left,
    this.right,
  });

  /// All borders with default style (single line, black).
  const DocxCellBorders.all({
    String color = '000000',
    int size = 4,
  })  : top = const DocxBorder(),
        bottom = const DocxBorder(),
        left = const DocxBorder(),
        right = const DocxBorder();

  /// No borders (explicitly removes borders from this cell).
  const DocxCellBorders.none()
      : top = null,
        bottom = null,
        left = null,
        right = null;

  /// Only bottom border (useful for underline-style separators).
  const DocxCellBorders.bottom({
    DocxBorder border = const DocxBorder(),
  })  : top = null,
        bottom = border,
        left = null,
        right = null;

  final DocxBorder? top;
  final DocxBorder? bottom;
  final DocxBorder? left;
  final DocxBorder? right;

  /// Returns true if any border is defined.
  bool get hasBorders =>
      top != null || bottom != null || left != null || right != null;
}

/// Padding inside a table cell, in twips (twentieths of a point).
///
/// A twip is 1/1440 of an inch, the unit Word stores cell margins in. Use
/// [DocxCellPadding.points] to declare the same thing in points:
///
/// ```dart
/// // 6pt top and bottom, 10pt left and right
/// const DocxCellPadding.points(top: 6, bottom: 6, left: 10, right: 10);
/// ```
///
/// Set it table-wide through [DocxTable.cellPadding] or per cell through
/// [DocxTableCell.padding]; a cell value overrides the table default.
class DocxCellPadding {
  /// Creates padding from raw twip values.
  const DocxCellPadding({
    this.top = 0,
    this.right = 0,
    this.bottom = 0,
    this.left = 0,
  });

  /// Creates uniform padding from a single twip value.
  const DocxCellPadding.all(int twips)
      : top = twips,
        right = twips,
        bottom = twips,
        left = twips;

  /// Creates padding from point values (1pt = 20 twips).
  const DocxCellPadding.points({
    int top = 0,
    int right = 0,
    int bottom = 0,
    int left = 0,
  })  : top = top * twipsPerPoint,
        right = right * twipsPerPoint,
        bottom = bottom * twipsPerPoint,
        left = left * twipsPerPoint;

  /// Word's own default cell margins: no vertical padding, 108 twips
  /// (0.075 inch) on each side.
  static const DocxCellPadding wordDefault =
      DocxCellPadding(left: 108, right: 108);

  /// Twips in a single point.
  static const int twipsPerPoint = 20;

  /// Padding above the cell content, in twips.
  final int top;

  /// Padding to the right of the cell content, in twips.
  final int right;

  /// Padding below the cell content, in twips.
  final int bottom;

  /// Padding to the left of the cell content, in twips.
  final int left;

  /// Whether any edge carries padding.
  bool get hasPadding => top > 0 || right > 0 || bottom > 0 || left > 0;

  @override
  bool operator ==(Object other) =>
      other is DocxCellPadding &&
      other.top == top &&
      other.right == right &&
      other.bottom == bottom &&
      other.left == left;

  @override
  int get hashCode => Object.hash(top, right, bottom, left);

  @override
  String toString() =>
      'DocxCellPadding(top: $top, right: $right, bottom: $bottom, left: $left)';
}

/// Represents a cell in a table row.
class DocxTableCell {
  const DocxTableCell({
    this.paragraphs = const [],
    this.alignment = DocxAlignment.left,
    this.verticalAlignment = DocxVerticalAlignment.top,
    this.backgroundColor,
    this.borders,
    this.colSpan = 1,
    this.rowSpan = 1,
    this.isMergedContinuation = false,
    this.padding,
  });

  /// Creates a simple cell with plain text.
  factory DocxTableCell.text(
    String text, {
    DocxAlignment alignment = DocxAlignment.left,
    DocxVerticalAlignment verticalAlignment = DocxVerticalAlignment.top,
    String? backgroundColor,
    DocxCellBorders? borders,
    int colSpan = 1,
    int rowSpan = 1,
    DocxCellPadding? padding,
  }) {
    return DocxTableCell(
      paragraphs: [DocxParagraph.text(text, alignment: alignment)],
      alignment: alignment,
      verticalAlignment: verticalAlignment,
      backgroundColor: backgroundColor,
      borders: borders,
      colSpan: colSpan,
      rowSpan: rowSpan,
      padding: padding,
    );
  }

  /// Creates a merged continuation cell (used for rowSpan > 1).
  /// This cell should be placed in subsequent rows where the merge continues.
  const factory DocxTableCell.merged() = _MergedCell;

  /// The paragraphs (content) in this cell.
  final List<DocxParagraph> paragraphs;

  /// Horizontal text alignment within the cell.
  final DocxAlignment alignment;

  /// Vertical alignment within the cell.
  final DocxVerticalAlignment verticalAlignment;

  /// Background color in hex format (e.g., "FFFF00" for yellow).
  final String? backgroundColor;

  /// Per-cell border configuration.
  /// When set, overrides the table-level borders for this cell.
  final DocxCellBorders? borders;

  /// Number of columns this cell spans (horizontal merge).
  /// Default is 1 (no merge).
  final int colSpan;

  /// Number of rows this cell spans (vertical merge).
  /// Default is 1 (no merge). Only set on the first cell of the merge.
  final int rowSpan;

  /// Whether this cell is a continuation of a vertical merge.
  /// Internal use only - use DocxTableCell.merged() factory instead.
  final bool isMergedContinuation;

  /// Padding inside this cell, overriding [DocxTable.cellPadding].
  ///
  /// Null falls back to the table-wide value.
  final DocxCellPadding? padding;
}

/// Internal class for merged continuation cells.
class _MergedCell implements DocxTableCell {
  const _MergedCell();

  @override
  List<DocxParagraph> get paragraphs => const [];

  @override
  DocxAlignment get alignment => DocxAlignment.left;

  @override
  DocxVerticalAlignment get verticalAlignment => DocxVerticalAlignment.top;

  @override
  String? get backgroundColor => null;

  @override
  DocxCellBorders? get borders => null;

  @override
  int get colSpan => 1;

  @override
  int get rowSpan => 1;

  @override
  bool get isMergedContinuation => true;

  @override
  DocxCellPadding? get padding => null;
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
    this.columnWidths,
    this.cellPadding,
  });

  /// Creates a simple table from a list of rows (list of cell texts).
  /// First row is treated as regular data (not styled as header).
  factory DocxTable.simple(
    List<List<String>> data, {
    DocxTableBorders borders = const DocxTableBorders.all(),
    List<double>? columnWidths,
    DocxCellPadding? cellPadding,
  }) {
    return DocxTable(
      rows: data
          .map((row) => DocxTableRow(
                cells: row.map((text) => DocxTableCell.text(text)).toList(),
              ))
          .toList(),
      borders: borders,
      columnWidths: columnWidths,
      cellPadding: cellPadding,
    );
  }

  /// Creates a table with styled header row.
  /// Headers get a background color and the data rows follow.
  factory DocxTable.fromHeaders({
    required List<String> headers,
    required List<List<String>> rows,
    String headerBackgroundColor = 'E0E0E0',
    DocxTableBorders borders = const DocxTableBorders.all(),
    List<double>? columnWidths,
    DocxCellPadding? cellPadding,
  }) {
    final headerRow = DocxTableRow(
      cells: headers
          .map((text) => DocxTableCell.text(
                text,
                backgroundColor: headerBackgroundColor,
              ))
          .toList(),
    );

    final dataRows = rows
        .map((row) => DocxTableRow(
              cells: row.map((text) => DocxTableCell.text(text)).toList(),
            ))
        .toList();

    return DocxTable(
      rows: [headerRow, ...dataRows],
      borders: borders,
      columnWidths: columnWidths,
      cellPadding: cellPadding,
    );
  }

  /// The rows in this table.
  final List<DocxTableRow> rows;

  /// Table border configuration.
  final DocxTableBorders borders;

  /// Column widths in percentages (e.g., [30, 40, 30] = 30%, 40%, 30%).
  /// If null, columns are evenly distributed.
  /// The values should sum to 100 for best results.
  final List<double>? columnWidths;

  /// Padding applied to every cell that does not set its own
  /// [DocxTableCell.padding].
  ///
  /// Null leaves the word processor's own defaults in place.
  final DocxCellPadding? cellPadding;

  /// Returns the number of columns (based on first row, accounting for colSpan).
  int get columnCount {
    if (rows.isEmpty) return 0;
    int count = 0;
    for (final cell in rows.first.cells) {
      count += cell.colSpan;
    }
    return count;
  }

  /// Returns the number of rows.
  int get rowCount => rows.length;
}

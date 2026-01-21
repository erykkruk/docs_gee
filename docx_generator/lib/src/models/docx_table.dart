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
    this.verticalAlignment = DocxVerticalAlignment.top,
    this.backgroundColor,
    this.colSpan = 1,
    this.rowSpan = 1,
    this.isMergedContinuation = false,
  });

  /// Creates a simple cell with plain text.
  factory DocxTableCell.text(
    String text, {
    DocxAlignment alignment = DocxAlignment.left,
    DocxVerticalAlignment verticalAlignment = DocxVerticalAlignment.top,
    String? backgroundColor,
    int colSpan = 1,
    int rowSpan = 1,
  }) {
    return DocxTableCell(
      paragraphs: [DocxParagraph.text(text, alignment: alignment)],
      alignment: alignment,
      verticalAlignment: verticalAlignment,
      backgroundColor: backgroundColor,
      colSpan: colSpan,
      rowSpan: rowSpan,
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

  /// Number of columns this cell spans (horizontal merge).
  /// Default is 1 (no merge).
  final int colSpan;

  /// Number of rows this cell spans (vertical merge).
  /// Default is 1 (no merge). Only set on the first cell of the merge.
  final int rowSpan;

  /// Whether this cell is a continuation of a vertical merge.
  /// Internal use only - use DocxTableCell.merged() factory instead.
  final bool isMergedContinuation;
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
  int get colSpan => 1;

  @override
  int get rowSpan => 1;

  @override
  bool get isMergedContinuation => true;
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
  });

  /// Creates a simple table from a list of rows (list of cell texts).
  /// First row is treated as regular data (not styled as header).
  factory DocxTable.simple(
    List<List<String>> data, {
    DocxTableBorders borders = const DocxTableBorders.all(),
    List<double>? columnWidths,
  }) {
    return DocxTable(
      rows: data
          .map((row) => DocxTableRow(
                cells: row.map((text) => DocxTableCell.text(text)).toList(),
              ))
          .toList(),
      borders: borders,
      columnWidths: columnWidths,
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

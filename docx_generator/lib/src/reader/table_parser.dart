import 'package:xml/xml.dart';

import '../models/models.dart';
import 'run_parser.dart';
import 'style_resolver.dart';

/// Parses `<w:tbl>` elements into [DocxTable] objects.
class TableParser {
  const TableParser._();

  /// Parses a `<w:tbl>` element into a [DocxTable].
  static DocxTable parse(
    XmlElement tblElement,
    Map<String, String> relationships,
  ) {
    final borders = _parseTableBorders(tblElement);
    final rows = <DocxTableRow>[];

    for (final child in tblElement.children) {
      if (child is! XmlElement || child.name.local != 'tr') continue;
      rows.add(_parseRow(child, relationships));
    }

    // Calculate rowSpan by analyzing vMerge patterns
    _resolveRowSpans(rows);

    return DocxTable(rows: rows, borders: borders);
  }

  /// Parses `<w:tblBorders>` from table properties.
  static DocxTableBorders _parseTableBorders(XmlElement tblElement) {
    final tblPr = _findChild(tblElement, 'tblPr');
    if (tblPr == null) return const DocxTableBorders.none();

    final tblBorders = _findChild(tblPr, 'tblBorders');
    if (tblBorders == null) return const DocxTableBorders.none();

    return DocxTableBorders(
      top: _parseBorder(tblBorders, 'top'),
      bottom: _parseBorder(tblBorders, 'bottom'),
      left: _parseBorder(tblBorders, 'left'),
      right: _parseBorder(tblBorders, 'right'),
      insideH: _parseBorder(tblBorders, 'insideH'),
      insideV: _parseBorder(tblBorders, 'insideV'),
    );
  }

  /// Parses a single border element (e.g. `<w:top w:val="single" w:sz="4" w:color="000000"/>`).
  static DocxBorder? _parseBorder(XmlElement parent, String name) {
    final element = _findChild(parent, name);
    if (element == null) return null;

    final val = _getAttr(element, 'val');
    if (val == 'nil' || val == 'none') return null;

    return DocxBorder(
      style: _parseBorderStyle(val),
      size: int.tryParse(_getAttr(element, 'sz') ?? '') ?? 4,
      color: _getAttr(element, 'color') ?? '000000',
    );
  }

  /// Parses border style string to [DocxBorderStyle].
  static DocxBorderStyle _parseBorderStyle(String? value) {
    return switch (value) {
      'single' => DocxBorderStyle.single,
      'double' => DocxBorderStyle.double,
      'dashed' => DocxBorderStyle.dashed,
      'dotted' => DocxBorderStyle.dotted,
      _ => DocxBorderStyle.single,
    };
  }

  /// Parses a `<w:tr>` element into a [DocxTableRow].
  static DocxTableRow _parseRow(
    XmlElement trElement,
    Map<String, String> relationships,
  ) {
    final cells = <DocxTableCell>[];

    for (final child in trElement.children) {
      if (child is! XmlElement || child.name.local != 'tc') continue;
      cells.add(_parseCell(child, relationships));
    }

    return DocxTableRow(cells: cells);
  }

  /// Parses a `<w:tc>` element into a [DocxTableCell].
  static DocxTableCell _parseCell(
    XmlElement tcElement,
    Map<String, String> relationships,
  ) {
    final tcPr = _findChild(tcElement, 'tcPr');

    // Check for vertical merge continuation
    final vMerge = _findChild(tcPr, 'vMerge');
    if (vMerge != null) {
      final vMergeVal = _getAttr(vMerge, 'val');
      if (vMergeVal == null || vMergeVal.isEmpty) {
        // Continuation cell (no val or empty val = continue previous merge)
        return const DocxTableCell.merged();
      }
      // val="restart" means this is the start of a new merge group
    }

    // ColSpan: <w:gridSpan w:val="3"/>
    final gridSpan = _findChild(tcPr, 'gridSpan');
    final colSpan = int.tryParse(_getAttr(gridSpan, 'val') ?? '') ?? 1;

    // Background color: <w:shd w:fill="E0E0E0"/>
    String? backgroundColor;
    final shd = _findChild(tcPr, 'shd');
    if (shd != null) {
      final fill = _getAttr(shd, 'fill');
      if (fill != null && fill != 'auto') {
        backgroundColor = fill;
      }
    }

    // Vertical alignment: <w:vAlign w:val="center"/>
    final vAlign = _findChild(tcPr, 'vAlign');
    final verticalAlignment =
        StyleResolver.resolveVerticalAlignment(_getAttr(vAlign, 'val'));

    // Cell borders: <w:tcBorders>
    final cellBorders = _parseCellBorders(tcPr);

    // Parse paragraphs inside the cell
    final paragraphs = <DocxParagraph>[];
    for (final child in tcElement.children) {
      if (child is! XmlElement || child.name.local != 'p') continue;
      paragraphs.add(_parseCellParagraph(child, relationships));
    }

    // Determine alignment from first paragraph (cell-level alignment)
    final alignment = paragraphs.isNotEmpty
        ? paragraphs.first.alignment
        : DocxAlignment.left;

    return DocxTableCell(
      paragraphs: paragraphs,
      alignment: alignment,
      verticalAlignment: verticalAlignment,
      backgroundColor: backgroundColor,
      borders: cellBorders,
      colSpan: colSpan,
      // rowSpan is resolved later in _resolveRowSpans
      rowSpan: vMerge != null ? 1 : 1,
    );
  }

  /// Parses `<w:tcBorders>` for cell-level border overrides.
  static DocxCellBorders? _parseCellBorders(XmlElement? tcPr) {
    if (tcPr == null) return null;
    final tcBorders = _findChild(tcPr, 'tcBorders');
    if (tcBorders == null) return null;

    final top = _parseBorder(tcBorders, 'top');
    final bottom = _parseBorder(tcBorders, 'bottom');
    final left = _parseBorder(tcBorders, 'left');
    final right = _parseBorder(tcBorders, 'right');

    if (top == null && bottom == null && left == null && right == null) {
      return null;
    }

    return DocxCellBorders(
      top: top,
      bottom: bottom,
      left: left,
      right: right,
    );
  }

  /// Parses a `<w:p>` element inside a table cell.
  static DocxParagraph _parseCellParagraph(
    XmlElement pElement,
    Map<String, String> relationships,
  ) {
    final pPr = _findChild(pElement, 'pPr');

    // Style
    final pStyle = _findChild(pPr, 'pStyle');
    final style = StyleResolver.resolveStyleId(_getAttr(pStyle, 'val'));

    // Alignment
    final jc = _findChild(pPr, 'jc');
    final alignment = StyleResolver.resolveAlignment(_getAttr(jc, 'val'));

    // Bookmark
    String? bookmarkName;
    final bookmarkStart = _findChild(pElement, 'bookmarkStart');
    if (bookmarkStart != null) {
      bookmarkName = _getAttr(bookmarkStart, 'name');
    }

    // Parse runs
    final runs = _parseRuns(pElement, relationships);

    return DocxParagraph(
      runs: runs,
      style: style,
      alignment: alignment,
      bookmarkName: bookmarkName,
    );
  }

  /// Parses all runs within a paragraph element, handling hyperlinks.
  static List<DocxRun> _parseRuns(
    XmlElement pElement,
    Map<String, String> relationships,
  ) {
    final runs = <DocxRun>[];

    for (final child in pElement.children) {
      if (child is! XmlElement) continue;

      if (child.name.local == 'r') {
        runs.add(RunParser.parse(child));
      } else if (child.name.local == 'hyperlink') {
        // External hyperlink: <w:hyperlink r:id="rId100">
        final rId = child.getAttribute('r:id') ??
            child.getAttribute('id',
                namespace:
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        final anchor = _getAttr(child, 'anchor');

        String? hyperlinkUrl;
        String? bookmarkRef;

        if (rId != null) {
          hyperlinkUrl = relationships[rId];
        } else if (anchor != null) {
          bookmarkRef = anchor;
        }

        for (final hChild in child.children) {
          if (hChild is XmlElement && hChild.name.local == 'r') {
            runs.add(RunParser.parse(
              hChild,
              hyperlink: hyperlinkUrl,
              bookmarkRef: bookmarkRef,
            ));
          }
        }
      }
    }

    return runs;
  }

  /// Resolves rowSpan by counting consecutive vMerge continuation cells.
  ///
  /// When a cell has `<w:vMerge w:val="restart"/>`, it starts a merge group.
  /// Subsequent rows with `<w:vMerge/>` (no val) at the same column index
  /// are continuation cells. The rowSpan of the restart cell equals the count.
  static void _resolveRowSpans(List<DocxTableRow> rows) {
    if (rows.isEmpty) return;

    final rowCount = rows.length;
    for (int rowIdx = 0; rowIdx < rowCount; rowIdx++) {
      final cells = rows[rowIdx].cells;
      for (int cellIdx = 0; cellIdx < cells.length; cellIdx++) {
        final cell = cells[cellIdx];
        if (cell.isMergedContinuation) continue;

        // Check if this cell starts a vertical merge
        // by looking at the next row's cell at the same index
        int span = 1;
        for (int nextRow = rowIdx + 1; nextRow < rowCount; nextRow++) {
          final nextCells = rows[nextRow].cells;
          if (cellIdx >= nextCells.length) break;
          if (!nextCells[cellIdx].isMergedContinuation) break;
          span++;
        }

        if (span > 1) {
          // Replace the cell with one that has the correct rowSpan
          cells[cellIdx] = DocxTableCell(
            paragraphs: cell.paragraphs,
            alignment: cell.alignment,
            verticalAlignment: cell.verticalAlignment,
            backgroundColor: cell.backgroundColor,
            borders: cell.borders,
            colSpan: cell.colSpan,
            rowSpan: span,
          );
        }
      }
    }
  }

  static XmlElement? _findChild(XmlElement? parent, String localName) {
    if (parent == null) return null;
    for (final child in parent.children) {
      if (child is XmlElement && child.name.local == localName) {
        return child;
      }
    }
    return null;
  }

  static String? _getAttr(XmlElement? element, String name) {
    if (element == null) return null;
    return element.getAttribute('w:$name') ?? element.getAttribute(name);
  }
}

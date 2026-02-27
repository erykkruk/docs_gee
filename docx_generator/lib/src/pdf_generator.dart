import 'dart:convert';
import 'dart:typed_data';

import 'document_generator.dart';
import 'models/models.dart';

/// Generates PDF files from [DocxDocument] without external dependencies.
///
/// Implements [DocumentGenerator] interface for interchangeable use with [DocxGenerator].
///
/// PDF is built manually by creating the required objects:
/// - Catalog (document root)
/// - Pages tree
/// - Page objects
/// - Font resources
/// - Content streams
class PdfGenerator implements DocumentGenerator {
  PdfGenerator({
    this.fontName = 'Helvetica',
    this.fontSize = 12,
    this.pageWidth = 612,
    this.pageHeight = 792,
    this.marginTop = 72,
    this.marginBottom = 72,
    this.marginLeft = 72,
    this.marginRight = 72,
  });

  /// Font name (PDF base 14 fonts: Helvetica, Times-Roman, Courier).
  final String fontName;

  /// Default font size in points.
  final int fontSize;

  /// Page width in points (612 = 8.5 inches).
  final double pageWidth;

  /// Page height in points (792 = 11 inches).
  final double pageHeight;

  /// Top margin in points.
  final double marginTop;

  /// Bottom margin in points.
  final double marginBottom;

  /// Left margin in points.
  final double marginLeft;

  /// Right margin in points.
  final double marginRight;

  /// Default file extension for PDF files.
  static const String defaultExtension = '.pdf';

  /// Generates a PDF document and returns it as bytes.
  @override
  Uint8List generate(DocxDocument document) {
    final builder = _PdfBuilder(
      fontName: fontName,
      fontSize: fontSize,
      pageWidth: pageWidth,
      pageHeight: pageHeight,
      marginTop: marginTop,
      marginBottom: marginBottom,
      marginLeft: marginLeft,
      marginRight: marginRight,
    );
    return builder.build(document);
  }
}

/// Internal PDF builder that handles the low-level PDF structure.
class _PdfBuilder {
  _PdfBuilder({
    required this.fontName,
    required this.fontSize,
    required this.pageWidth,
    required this.pageHeight,
    required this.marginTop,
    required this.marginBottom,
    required this.marginLeft,
    required this.marginRight,
  });

  final String fontName;
  final int fontSize;
  final double pageWidth;
  final double pageHeight;
  final double marginTop;
  final double marginBottom;
  final double marginLeft;
  final double marginRight;

  final List<_PdfObject> _objects = [];
  final List<int> _objectOffsets = [];
  int _currentOffset = 0;

  double get _contentWidth => pageWidth - marginLeft - marginRight;
  double get _contentHeight => pageHeight - marginTop - marginBottom;

  Uint8List build(DocxDocument document) {
    final buffer = BytesBuilder();

    // PDF Header
    _writeBytes(buffer, '%PDF-1.4\n');
    // Binary marker (recommended for binary content)
    _writeBytes(buffer, '%\xE2\xE3\xCF\xD3\n');

    // Split document into pages
    final pages = _paginateDocument(document);

    // Object 1: Catalog
    final catalogObj = _PdfObject(1, _buildCatalog());
    _objects.add(catalogObj);

    // Object 2: Pages
    final pagesObj = _PdfObject(2, _buildPages(pages.length));
    _objects.add(pagesObj);

    // Object 3: Font
    final fontObj = _PdfObject(3, _buildFont());
    _objects.add(fontObj);

    // Object 4: Bold Font
    final boldFontObj = _PdfObject(4, _buildBoldFont());
    _objects.add(boldFontObj);

    // Object 5: Italic Font
    final italicFontObj = _PdfObject(5, _buildItalicFont());
    _objects.add(italicFontObj);

    // Object 6: Bold-Italic Font
    final boldItalicFontObj = _PdfObject(6, _buildBoldItalicFont());
    _objects.add(boldItalicFontObj);

    // Object 7: Monospace Font (Courier)
    final monoFontObj = _PdfObject(7, _buildMonoFont());
    _objects.add(monoFontObj);

    // Object 8: Info dictionary (metadata)
    final infoObj = _PdfObject(8, _buildInfo(document));
    _objects.add(infoObj);

    // Create page objects and content streams
    int nextObjId = 9;
    for (int i = 0; i < pages.length; i++) {
      final pageObjId = nextObjId++;
      final contentObjId = nextObjId++;

      final contentStream = _buildContentStream(pages[i]);
      final contentObj =
          _PdfObject(contentObjId, _buildStreamObject(contentStream));

      final pageObj = _PdfObject(pageObjId, _buildPage(contentObjId));

      _objects.add(pageObj);
      _objects.add(contentObj);
    }

    // Write all objects
    for (final obj in _objects) {
      _objectOffsets.add(_currentOffset);
      final objStr = '${obj.id} 0 obj\n${obj.content}\nendobj\n';
      _writeBytes(buffer, objStr);
    }

    // Write xref table
    final xrefOffset = _currentOffset;
    _writeBytes(buffer, 'xref\n');
    _writeBytes(buffer, '0 ${_objects.length + 1}\n');
    _writeBytes(buffer, '0000000000 65535 f \n');
    for (final offset in _objectOffsets) {
      _writeBytes(buffer, '${offset.toString().padLeft(10, '0')} 00000 n \n');
    }

    // Write trailer
    _writeBytes(buffer, 'trailer\n');
    _writeBytes(buffer, '<<\n');
    _writeBytes(buffer, '/Size ${_objects.length + 1}\n');
    _writeBytes(buffer, '/Root 1 0 R\n');
    _writeBytes(buffer, '/Info 8 0 R\n');
    _writeBytes(buffer, '>>\n');
    _writeBytes(buffer, 'startxref\n');
    _writeBytes(buffer, '$xrefOffset\n');
    _writeBytes(buffer, '%%EOF\n');

    return buffer.toBytes();
  }

  void _writeBytes(BytesBuilder buffer, String str) {
    final bytes = utf8.encode(str);
    buffer.add(bytes);
    _currentOffset += bytes.length;
  }

  String _buildCatalog() {
    return '<<\n/Type /Catalog\n/Pages 2 0 R\n>>';
  }

  String _buildPages(int pageCount) {
    final kids = StringBuffer();
    kids.write('/Kids [');
    // Page objects start at ID 9, every other object (page, content, page, content...)
    for (int i = 0; i < pageCount; i++) {
      final pageObjId = 9 + (i * 2);
      kids.write('$pageObjId 0 R ');
    }
    kids.write(']');

    return '<<\n'
        '/Type /Pages\n'
        '${kids.toString()}\n'
        '/Count $pageCount\n'
        '>>';
  }

  String _buildFont() {
    return '<<\n'
        '/Type /Font\n'
        '/Subtype /Type1\n'
        '/BaseFont /$fontName\n'
        '/Encoding /WinAnsiEncoding\n'
        '>>';
  }

  String _buildBoldFont() {
    final boldName = _getBoldFontName();
    return '<<\n'
        '/Type /Font\n'
        '/Subtype /Type1\n'
        '/BaseFont /$boldName\n'
        '/Encoding /WinAnsiEncoding\n'
        '>>';
  }

  String _buildItalicFont() {
    final italicName = _getItalicFontName();
    return '<<\n'
        '/Type /Font\n'
        '/Subtype /Type1\n'
        '/BaseFont /$italicName\n'
        '/Encoding /WinAnsiEncoding\n'
        '>>';
  }

  String _buildBoldItalicFont() {
    final boldItalicName = _getBoldItalicFontName();
    return '<<\n'
        '/Type /Font\n'
        '/Subtype /Type1\n'
        '/BaseFont /$boldItalicName\n'
        '/Encoding /WinAnsiEncoding\n'
        '>>';
  }

  String _getBoldFontName() {
    if (fontName == 'Helvetica') return 'Helvetica-Bold';
    if (fontName == 'Times-Roman') return 'Times-Bold';
    if (fontName == 'Courier') return 'Courier-Bold';
    return '$fontName-Bold';
  }

  String _getItalicFontName() {
    if (fontName == 'Helvetica') return 'Helvetica-Oblique';
    if (fontName == 'Times-Roman') return 'Times-Italic';
    if (fontName == 'Courier') return 'Courier-Oblique';
    return '$fontName-Italic';
  }

  String _getBoldItalicFontName() {
    if (fontName == 'Helvetica') return 'Helvetica-BoldOblique';
    if (fontName == 'Times-Roman') return 'Times-BoldItalic';
    if (fontName == 'Courier') return 'Courier-BoldOblique';
    return '$fontName-BoldItalic';
  }

  String _buildMonoFont() {
    return '<<\n'
        '/Type /Font\n'
        '/Subtype /Type1\n'
        '/BaseFont /Courier\n'
        '/Encoding /WinAnsiEncoding\n'
        '>>';
  }

  String _buildInfo(DocxDocument document) {
    final buffer = StringBuffer();
    buffer.writeln('<<');
    if (document.title != null) {
      buffer.writeln('/Title (${_escapePdfString(document.title!)})');
    }
    if (document.author != null) {
      buffer.writeln('/Author (${_escapePdfString(document.author!)})');
    }
    buffer.writeln('/Creator (docs_gee)');
    buffer.writeln('/Producer (docs_gee Dart library)');
    // Add creation date in PDF format: D:YYYYMMDDHHmmss
    final now = DateTime.now();
    final dateStr = 'D:${now.year.toString().padLeft(4, '0')}'
        '${now.month.toString().padLeft(2, '0')}'
        '${now.day.toString().padLeft(2, '0')}'
        '${now.hour.toString().padLeft(2, '0')}'
        '${now.minute.toString().padLeft(2, '0')}'
        '${now.second.toString().padLeft(2, '0')}';
    buffer.writeln('/CreationDate ($dateStr)');
    buffer.writeln('/ModDate ($dateStr)');
    buffer.write('>>');
    return buffer.toString();
  }

  String _buildPage(int contentObjId) {
    return '<<\n'
        '/Type /Page\n'
        '/Parent 2 0 R\n'
        '/MediaBox [0 0 $pageWidth $pageHeight]\n'
        '/Contents $contentObjId 0 R\n'
        '/Resources <<\n'
        '  /Font <<\n'
        '    /F1 3 0 R\n'
        '    /F2 4 0 R\n'
        '    /F3 5 0 R\n'
        '    /F4 6 0 R\n'
        '    /F5 7 0 R\n'
        '  >>\n'
        '>>\n'
        '>>';
  }

  String _buildStreamObject(String content) {
    final bytes = utf8.encode(content);
    return '<<\n/Length ${bytes.length}\n>>\nstream\n$content\nendstream';
  }

  List<List<Object>> _paginateDocument(DocxDocument document) {
    final pages = <List<Object>>[];
    var currentPage = <Object>[];
    var currentY = _contentHeight;

    for (final item in document.content) {
      double itemHeight;
      bool hasPageBreak = false;

      if (item is DocxParagraph) {
        itemHeight = _estimateParagraphHeight(item);
        hasPageBreak = item.pageBreakBefore;
      } else if (item is DocxTable) {
        itemHeight = _estimateTableHeight(item);
      } else {
        continue;
      }

      // Check for explicit page break
      if (hasPageBreak && currentPage.isNotEmpty) {
        pages.add(currentPage);
        currentPage = <Object>[];
        currentY = _contentHeight;
      }

      // Check if item fits on current page
      if (currentY - itemHeight < 0 && currentPage.isNotEmpty) {
        pages.add(currentPage);
        currentPage = <Object>[];
        currentY = _contentHeight;
      }

      currentPage.add(item);
      currentY -= itemHeight;
    }

    if (currentPage.isNotEmpty) {
      pages.add(currentPage);
    }

    // Ensure at least one empty page if document is empty
    if (pages.isEmpty) {
      pages.add([]);
    }

    return pages;
  }

  /// Estimates the height of a table for pagination.
  double _estimateTableHeight(DocxTable table) {
    double totalHeight = 0;
    for (final row in table.rows) {
      totalHeight += _estimateRowHeight(row);
    }
    return totalHeight + 10; // Add some padding
  }

  /// Estimates the height of a table row.
  double _estimateRowHeight(DocxTableRow row) {
    double maxHeight = fontSize * 1.5; // Minimum row height
    for (final cell in row.cells) {
      double cellHeight = 0;
      for (final paragraph in cell.paragraphs) {
        cellHeight += _estimateParagraphHeight(paragraph);
      }
      if (cellHeight > maxHeight) maxHeight = cellHeight;
    }
    return maxHeight + 8; // Add cell padding
  }

  double _estimateParagraphHeight(DocxParagraph paragraph) {
    final size = _getFontSizeForStyle(paragraph.style);
    final lineHeight = size * 1.5;
    final text = paragraph.plainText;

    // Estimate number of lines based on character count and content width
    // Approximate: 6 points per character for standard fonts
    final charsPerLine = (_contentWidth / (size * 0.5)).floor();
    final lineCount = (text.length / charsPerLine).ceil().clamp(1, 100);

    return lineCount * lineHeight + (size * 0.5); // Add spacing after paragraph
  }

  int _getFontSizeForStyle(DocxParagraphStyle style) {
    return switch (style) {
      DocxParagraphStyle.heading1 => (fontSize * 2).round(),
      DocxParagraphStyle.heading2 => (fontSize * 1.5).round(),
      DocxParagraphStyle.heading3 => (fontSize * 1.25).round(),
      DocxParagraphStyle.heading4 => (fontSize * 1.1).round(),
      DocxParagraphStyle.subtitle => (fontSize * 1.2).round(),
      DocxParagraphStyle.caption => (fontSize * 0.85).round(),
      DocxParagraphStyle.codeBlock => (fontSize * 0.85).round(),
      DocxParagraphStyle.footnote => (fontSize * 0.75).round(),
      DocxParagraphStyle.quote => fontSize,
      _ => fontSize,
    };
  }

  String _buildContentStream(List<Object> contentItems) {
    final buffer = StringBuffer();
    var currentY = pageHeight - marginTop;

    // Track counters for each list type and level
    final listCounters = <String, int>{};
    DocxParagraphStyle? lastListStyle;
    int lastIndentLevel = -1;

    for (final item in contentItems) {
      if (item is DocxParagraph) {
        currentY = _renderParagraph(
          buffer,
          item,
          currentY,
          listCounters,
          lastListStyle,
          lastIndentLevel,
        );
        if (item.style.isList) {
          lastListStyle = item.style;
          lastIndentLevel = item.indentLevel;
        } else {
          listCounters.clear();
          lastListStyle = null;
          lastIndentLevel = -1;
        }
      } else if (item is DocxTable) {
        currentY = _renderTable(buffer, item, currentY);
      }
    }

    return buffer.toString();
  }

  /// Renders a paragraph and returns the new Y position.
  double _renderParagraph(
    StringBuffer buffer,
    DocxParagraph paragraph,
    double startY,
    Map<String, int> listCounters,
    DocxParagraphStyle? lastListStyle,
    int lastIndentLevel,
  ) {
    buffer.writeln('BT'); // Begin text

    var currentY = startY;
    final size = _getFontSizeForStyle(paragraph.style);
    final lineHeight = size * 1.4;

    // Handle list prefixes and indentation
    String prefix = '';
    final indentOffset = paragraph.indentLevel * 18.0;

    if (paragraph.style.isList) {
      // Reset counters if switching list type at same level
      if (lastListStyle != paragraph.style ||
          paragraph.indentLevel < lastIndentLevel) {
        listCounters.removeWhere((key, _) {
          final parts = key.split('_');
          if (parts.length == 2) {
            final level = int.tryParse(parts[1]) ?? 0;
            return level >= paragraph.indentLevel;
          }
          return false;
        });
      }

      final counterKey = '${paragraph.style.name}_${paragraph.indentLevel}';

      if (paragraph.style == DocxParagraphStyle.listBullet) {
        prefix = '  \u2022 '; // Unicode bullet, converted in _escapePdfString
      } else if (paragraph.style == DocxParagraphStyle.listDash) {
        prefix = '  - ';
      } else if (paragraph.style == DocxParagraphStyle.listNumber) {
        final count = (listCounters[counterKey] ?? 0) + 1;
        listCounters[counterKey] = count;
        prefix = '  $count. ';
      } else if (paragraph.style == DocxParagraphStyle.listNumberAlpha) {
        final count = (listCounters[counterKey] ?? 0) + 1;
        listCounters[counterKey] = count;
        prefix = '  ${_toAlpha(count)}) ';
      } else if (paragraph.style == DocxParagraphStyle.listNumberRoman) {
        final count = (listCounters[counterKey] ?? 0) + 1;
        listCounters[counterKey] = count;
        prefix = '  ${_toRoman(count)}. ';
      }
    }

    // Combine all runs into text segments
    final segments = <_TextSegment>[];
    final isHeading = paragraph.style == DocxParagraphStyle.heading1 ||
        paragraph.style == DocxParagraphStyle.heading2 ||
        paragraph.style == DocxParagraphStyle.heading3 ||
        paragraph.style == DocxParagraphStyle.heading4;
    final isItalicStyle = paragraph.style == DocxParagraphStyle.quote ||
        paragraph.style == DocxParagraphStyle.subtitle;
    final isCodeBlock = paragraph.style == DocxParagraphStyle.codeBlock;
    final isFootnote = paragraph.style == DocxParagraphStyle.footnote;

    if (prefix.isNotEmpty) {
      segments.add(_TextSegment(prefix, '/F1'));
    }

    for (final run in paragraph.runs) {
      String fontRef;
      String? textColor = run.color;

      if (isCodeBlock) {
        fontRef = '/F5';
      } else {
        final isBold = run.bold || isHeading;
        final isItalic = run.italic || isItalicStyle;
        if (isBold && isItalic) {
          fontRef = '/F4';
        } else if (isBold) {
          fontRef = '/F2';
        } else if (isItalic) {
          fontRef = '/F3';
        } else {
          fontRef = '/F1';
        }
      }

      if (isFootnote && textColor == null) {
        textColor = '666666';
      }

      segments.add(_TextSegment(
        run.text,
        fontRef,
        color: textColor,
        underline: run.underline,
        strikethrough: run.strikethrough,
      ));
    }

    final lines = _wrapTextSegments(segments, size);

    for (int i = 0; i < lines.length; i++) {
      final line = lines[i];
      final lineWidth = line.fold<double>(
        0,
        (sum, s) => sum + _estimateTextWidth(s.text, size, fontRef: s.fontRef),
      );

      double xPos;
      switch (paragraph.alignment) {
        case DocxAlignment.center:
          xPos = marginLeft +
              indentOffset +
              (_contentWidth - indentOffset - lineWidth) / 2;
        case DocxAlignment.right:
          xPos = marginLeft + _contentWidth - lineWidth;
        case DocxAlignment.justify:
        case DocxAlignment.left:
          xPos = marginLeft + indentOffset;
      }

      buffer.writeln('1 0 0 1 $xPos $currentY Tm');
      var segmentX = xPos;

      for (final segment in line) {
        if (segment.color != null) {
          final rgb = _hexToRgb(segment.color!);
          buffer.writeln('${rgb.$1} ${rgb.$2} ${rgb.$3} rg');
        } else {
          buffer.writeln('0 0 0 rg');
        }

        buffer.writeln('${segment.fontRef} $size Tf');
        buffer.writeln('(${_escapePdfString(segment.text)}) Tj');

        final segmentWidth =
            _estimateTextWidth(segment.text, size, fontRef: segment.fontRef);

        if (segment.underline || segment.strikethrough) {
          buffer.writeln('ET');
          buffer.writeln('q');

          if (segment.color != null) {
            final rgb = _hexToRgb(segment.color!);
            buffer.writeln('${rgb.$1} ${rgb.$2} ${rgb.$3} RG');
          } else {
            buffer.writeln('0 0 0 RG');
          }

          buffer.writeln('0.5 w');

          if (segment.underline) {
            final underlineY = currentY - 2;
            buffer.writeln('$segmentX $underlineY m');
            buffer.writeln('${segmentX + segmentWidth} $underlineY l');
            buffer.writeln('S');
          }

          if (segment.strikethrough) {
            final strikeY = currentY + size * 0.3;
            buffer.writeln('$segmentX $strikeY m');
            buffer.writeln('${segmentX + segmentWidth} $strikeY l');
            buffer.writeln('S');
          }

          buffer.writeln('Q');
          buffer.writeln('BT');
          buffer.writeln('1 0 0 1 ${segmentX + segmentWidth} $currentY Tm');
        }

        segmentX += segmentWidth;
      }

      currentY -= lineHeight;
    }

    buffer.writeln('ET');
    currentY -= size * 0.4;

    return currentY;
  }

  /// Calculates column widths for PDF based on table configuration.
  List<double> _calculatePdfColumnWidths(DocxTable table) {
    final colCount = table.columnCount;
    if (colCount == 0) return [];

    if (table.columnWidths != null && table.columnWidths!.length == colCount) {
      // Use custom widths (percentages converted to points)
      return table.columnWidths!
          .map((percent) => _contentWidth * percent / 100)
          .toList();
    } else {
      // Even distribution
      final width = _contentWidth / colCount;
      return List.filled(colCount, width);
    }
  }

  /// Renders a table and returns the new Y position.
  double _renderTable(StringBuffer buffer, DocxTable table, double startY) {
    if (table.rows.isEmpty || table.columnCount == 0) return startY;

    var currentY = startY;
    final columnWidths = _calculatePdfColumnWidths(table);

    for (int rowIndex = 0; rowIndex < table.rows.length; rowIndex++) {
      final row = table.rows[rowIndex];
      final rowHeight = _estimateRowHeight(row);
      var currentX = marginLeft;
      int gridColIndex = 0;

      for (int cellIndex = 0; cellIndex < row.cells.length; cellIndex++) {
        final cell = row.cells[cellIndex];

        // Skip merged continuation cells (they don't render content)
        if (cell.isMergedContinuation) {
          if (gridColIndex < columnWidths.length) {
            currentX += columnWidths[gridColIndex];
            gridColIndex++;
          }
          continue;
        }

        // Calculate cell width based on colspan
        double cellWidth = 0;
        for (int i = 0;
            i < cell.colSpan && gridColIndex + i < columnWidths.length;
            i++) {
          cellWidth += columnWidths[gridColIndex + i];
        }
        if (cellWidth == 0 && columnWidths.isNotEmpty) {
          cellWidth = columnWidths.first;
        }

        // Draw cell background
        if (cell.backgroundColor != null) {
          final rgb = _hexToRgb(cell.backgroundColor!);
          buffer.writeln('q');
          buffer.writeln('${rgb.$1} ${rgb.$2} ${rgb.$3} rg');
          buffer.writeln(
              '$currentX ${currentY - rowHeight} $cellWidth $rowHeight re');
          buffer.writeln('f');
          buffer.writeln('Q');
        }

        // Draw cell borders
        _drawCellBorders(
          buffer,
          table.borders,
          currentX,
          currentY,
          cellWidth,
          rowHeight,
          isFirstRow: rowIndex == 0,
          isLastRow: rowIndex == table.rows.length - 1,
          isFirstCol: gridColIndex == 0,
          isLastCol: gridColIndex + cell.colSpan >= columnWidths.length,
          cellBorders: cell.borders,
        );

        // Render cell content with vertical alignment
        const cellPadding = 4.0;
        final contentHeight = _estimateCellContentHeight(cell);
        double cellY;

        switch (cell.verticalAlignment) {
          case DocxVerticalAlignment.center:
            cellY = currentY - cellPadding - (rowHeight - contentHeight) / 2;
          case DocxVerticalAlignment.bottom:
            cellY = currentY - rowHeight + contentHeight;
          case DocxVerticalAlignment.top:
            cellY = currentY - cellPadding;
        }

        for (final paragraph in cell.paragraphs) {
          cellY = _renderCellParagraph(
            buffer,
            paragraph,
            cellY,
            currentX + cellPadding,
            cellWidth - (cellPadding * 2),
          );
        }

        currentX += cellWidth;
        gridColIndex += cell.colSpan;
      }

      currentY -= rowHeight;
    }

    return currentY - 10; // Add spacing after table
  }

  /// Estimates the content height of a cell for vertical alignment.
  double _estimateCellContentHeight(DocxTableCell cell) {
    double height = 0;
    for (final paragraph in cell.paragraphs) {
      final size = _getFontSizeForStyle(paragraph.style);
      final lineHeight = size * 1.4;
      // Rough estimate: one line per paragraph
      height += lineHeight;
    }
    return height;
  }

  /// Draws cell borders.
  ///
  /// If [cellBorders] is provided, it overrides the table-level borders
  /// for this specific cell.
  void _drawCellBorders(
    StringBuffer buffer,
    DocxTableBorders borders,
    double x,
    double y,
    double width,
    double height, {
    required bool isFirstRow,
    required bool isLastRow,
    required bool isFirstCol,
    required bool isLastCol,
    DocxCellBorders? cellBorders,
  }) {
    // Determine effective borders for each side
    final bool drawTop;
    final bool drawBottom;
    final bool drawLeft;
    final bool drawRight;

    DocxBorder? topBorder;
    DocxBorder? bottomBorder;
    DocxBorder? leftBorder;
    DocxBorder? rightBorder;

    if (cellBorders != null) {
      // Cell-level borders override table-level
      drawTop = cellBorders.top != null;
      drawBottom = cellBorders.bottom != null;
      drawLeft = cellBorders.left != null;
      drawRight = cellBorders.right != null;
      topBorder = cellBorders.top;
      bottomBorder = cellBorders.bottom;
      leftBorder = cellBorders.left;
      rightBorder = cellBorders.right;
    } else {
      // Fall back to table-level borders
      drawTop = (isFirstRow && borders.top != null) ||
          (!isFirstRow && borders.insideH != null);
      drawBottom = isLastRow && borders.bottom != null;
      drawLeft = (isFirstCol && borders.left != null) ||
          (!isFirstCol && borders.insideV != null);
      drawRight = isLastCol && borders.right != null;
      topBorder = isFirstRow ? borders.top : borders.insideH;
      bottomBorder = borders.bottom;
      leftBorder = isFirstCol ? borders.left : borders.insideV;
      rightBorder = borders.right;
    }

    if (!drawTop && !drawBottom && !drawLeft && !drawRight) return;

    buffer.writeln('q');

    // Top border
    if (drawTop && topBorder != null) {
      final rgb = _hexToRgb(topBorder.color);
      buffer.writeln('${rgb.$1} ${rgb.$2} ${rgb.$3} RG');
      buffer.writeln('${topBorder.size / 8} w');
      buffer.writeln('$x $y m');
      buffer.writeln('${x + width} $y l');
      buffer.writeln('S');
    }

    // Bottom border
    if (drawBottom && bottomBorder != null) {
      final bottomY = y - height;
      final rgb = _hexToRgb(bottomBorder.color);
      buffer.writeln('${rgb.$1} ${rgb.$2} ${rgb.$3} RG');
      buffer.writeln('${bottomBorder.size / 8} w');
      buffer.writeln('$x $bottomY m');
      buffer.writeln('${x + width} $bottomY l');
      buffer.writeln('S');
    }

    // Left border
    if (drawLeft && leftBorder != null) {
      final rgb = _hexToRgb(leftBorder.color);
      buffer.writeln('${rgb.$1} ${rgb.$2} ${rgb.$3} RG');
      buffer.writeln('${leftBorder.size / 8} w');
      buffer.writeln('$x $y m');
      buffer.writeln('$x ${y - height} l');
      buffer.writeln('S');
    }

    // Right border
    if (drawRight && rightBorder != null) {
      final rightX = x + width;
      final rgb = _hexToRgb(rightBorder.color);
      buffer.writeln('${rgb.$1} ${rgb.$2} ${rgb.$3} RG');
      buffer.writeln('${rightBorder.size / 8} w');
      buffer.writeln('$rightX $y m');
      buffer.writeln('$rightX ${y - height} l');
      buffer.writeln('S');
    }

    buffer.writeln('Q');
  }

  /// Renders a paragraph inside a cell and returns the new Y position.
  double _renderCellParagraph(
    StringBuffer buffer,
    DocxParagraph paragraph,
    double startY,
    double cellX,
    double cellWidth,
  ) {
    buffer.writeln('BT');

    var currentY = startY;
    final size = _getFontSizeForStyle(paragraph.style);
    final lineHeight = size * 1.4;

    final segments = <_TextSegment>[];
    final isHeading = paragraph.style == DocxParagraphStyle.heading1 ||
        paragraph.style == DocxParagraphStyle.heading2 ||
        paragraph.style == DocxParagraphStyle.heading3 ||
        paragraph.style == DocxParagraphStyle.heading4;
    final isItalicStyle = paragraph.style == DocxParagraphStyle.quote ||
        paragraph.style == DocxParagraphStyle.subtitle;
    final isCodeBlock = paragraph.style == DocxParagraphStyle.codeBlock;
    final isFootnote = paragraph.style == DocxParagraphStyle.footnote;

    for (final run in paragraph.runs) {
      String fontRef;
      String? textColor = run.color;

      if (isCodeBlock) {
        fontRef = '/F5';
      } else {
        final isBold = run.bold || isHeading;
        final isItalic = run.italic || isItalicStyle;
        if (isBold && isItalic) {
          fontRef = '/F4';
        } else if (isBold) {
          fontRef = '/F2';
        } else if (isItalic) {
          fontRef = '/F3';
        } else {
          fontRef = '/F1';
        }
      }

      if (isFootnote && textColor == null) {
        textColor = '666666';
      }

      segments.add(_TextSegment(
        run.text,
        fontRef,
        color: textColor,
        underline: run.underline,
        strikethrough: run.strikethrough,
      ));
    }

    final lines = _wrapTextSegmentsForWidth(segments, size, cellWidth);

    for (final line in lines) {
      final lineWidth = line.fold<double>(
        0,
        (sum, s) => sum + _estimateTextWidth(s.text, size, fontRef: s.fontRef),
      );

      double xPos;
      switch (paragraph.alignment) {
        case DocxAlignment.center:
          xPos = cellX + (cellWidth - lineWidth) / 2;
        case DocxAlignment.right:
          xPos = cellX + cellWidth - lineWidth;
        case DocxAlignment.justify:
        case DocxAlignment.left:
          xPos = cellX;
      }

      buffer.writeln('1 0 0 1 $xPos $currentY Tm');

      var segmentX = xPos;

      for (final segment in line) {
        if (segment.color != null) {
          final rgb = _hexToRgb(segment.color!);
          buffer.writeln('${rgb.$1} ${rgb.$2} ${rgb.$3} rg');
        } else {
          buffer.writeln('0 0 0 rg');
        }

        buffer.writeln('${segment.fontRef} $size Tf');
        buffer.writeln('(${_escapePdfString(segment.text)}) Tj');

        final segmentWidth =
            _estimateTextWidth(segment.text, size, fontRef: segment.fontRef);

        if (segment.underline || segment.strikethrough) {
          buffer.writeln('ET');
          buffer.writeln('q');

          if (segment.color != null) {
            final rgb = _hexToRgb(segment.color!);
            buffer.writeln('${rgb.$1} ${rgb.$2} ${rgb.$3} RG');
          } else {
            buffer.writeln('0 0 0 RG');
          }

          buffer.writeln('0.5 w');

          if (segment.underline) {
            final underlineY = currentY - 2;
            buffer.writeln('$segmentX $underlineY m');
            buffer.writeln('${segmentX + segmentWidth} $underlineY l');
            buffer.writeln('S');
          }

          if (segment.strikethrough) {
            final strikeY = currentY + size * 0.3;
            buffer.writeln('$segmentX $strikeY m');
            buffer.writeln('${segmentX + segmentWidth} $strikeY l');
            buffer.writeln('S');
          }

          buffer.writeln('Q');
          buffer.writeln('BT');
          buffer.writeln('1 0 0 1 ${segmentX + segmentWidth} $currentY Tm');
        }

        segmentX += segmentWidth;
      }

      currentY -= lineHeight;
    }

    buffer.writeln('ET');
    return currentY;
  }

  /// Wraps text segments for a specific width (used in table cells).
  List<List<_TextSegment>> _wrapTextSegmentsForWidth(
      List<_TextSegment> segments, int size, double maxWidth) {
    final lines = <List<_TextSegment>>[];
    var currentLine = <_TextSegment>[];
    var currentLineWidth = 0.0;
    for (final segment in segments) {
      final spaceWidth =
          _estimateTextWidth(' ', size, fontRef: segment.fontRef);
      final words = segment.text.split(' ');

      for (int i = 0; i < words.length; i++) {
        final word = words[i];
        if (word.isEmpty) continue;

        final wordWidth =
            _estimateTextWidth(word, size, fontRef: segment.fontRef);
        final needsSpace = currentLine.isNotEmpty &&
            (currentLine.last.text.isNotEmpty &&
                !currentLine.last.text.endsWith(' '));
        final additionalWidth = needsSpace ? spaceWidth + wordWidth : wordWidth;

        if (currentLineWidth + additionalWidth > maxWidth &&
            currentLine.isNotEmpty) {
          lines.add(currentLine);
          currentLine = <_TextSegment>[];
          currentLineWidth = 0.0;

          currentLine.add(_TextSegment(
            word,
            segment.fontRef,
            color: segment.color,
            underline: segment.underline,
            strikethrough: segment.strikethrough,
          ));
          currentLineWidth = wordWidth;
        } else {
          if (needsSpace) {
            currentLine.add(_TextSegment(
              ' $word',
              segment.fontRef,
              color: segment.color,
              underline: segment.underline,
              strikethrough: segment.strikethrough,
            ));
            currentLineWidth += spaceWidth + wordWidth;
          } else {
            currentLine.add(_TextSegment(
              word,
              segment.fontRef,
              color: segment.color,
              underline: segment.underline,
              strikethrough: segment.strikethrough,
            ));
            currentLineWidth += wordWidth;
          }
        }
      }
    }

    if (currentLine.isNotEmpty) {
      lines.add(currentLine);
    }

    if (lines.isEmpty) {
      lines.add([_TextSegment('', '/F1')]);
    }

    return lines;
  }

  /// Converts hex color string to RGB values (0-1 range).
  (double, double, double) _hexToRgb(String hex) {
    final cleanHex = hex.replaceAll('#', '');
    if (cleanHex.length != 6) return (0.0, 0.0, 0.0);

    final r = int.parse(cleanHex.substring(0, 2), radix: 16) / 255;
    final g = int.parse(cleanHex.substring(2, 4), radix: 16) / 255;
    final b = int.parse(cleanHex.substring(4, 6), radix: 16) / 255;
    return (r, g, b);
  }

  /// Converts number to lowercase alphabetic (1=a, 2=b, ..., 26=z, 27=aa, etc.)
  String _toAlpha(int n) {
    if (n <= 0) return '';
    String result = '';
    int num = n;
    while (num > 0) {
      num--;
      result = String.fromCharCode('a'.codeUnitAt(0) + (num % 26)) + result;
      num ~/= 26;
    }
    return result;
  }

  /// Converts number to uppercase Roman numerals.
  String _toRoman(int n) {
    if (n <= 0 || n > 3999) return n.toString();

    const romanNumerals = [
      (1000, 'M'),
      (900, 'CM'),
      (500, 'D'),
      (400, 'CD'),
      (100, 'C'),
      (90, 'XC'),
      (50, 'L'),
      (40, 'XL'),
      (10, 'X'),
      (9, 'IX'),
      (5, 'V'),
      (4, 'IV'),
      (1, 'I'),
    ];

    final buffer = StringBuffer();
    int remaining = n;
    for (final (value, numeral) in romanNumerals) {
      while (remaining >= value) {
        buffer.write(numeral);
        remaining -= value;
      }
    }
    return buffer.toString();
  }

  /// Wraps text segments into lines that fit within content width.
  List<List<_TextSegment>> _wrapTextSegments(
      List<_TextSegment> segments, int size) {
    final lines = <List<_TextSegment>>[];
    var currentLine = <_TextSegment>[];
    var currentLineWidth = 0.0;
    final maxWidth = _contentWidth;

    for (final segment in segments) {
      final spaceWidth =
          _estimateTextWidth(' ', size, fontRef: segment.fontRef);
      final words = segment.text.split(' ');

      for (int i = 0; i < words.length; i++) {
        final word = words[i];
        if (word.isEmpty) continue;

        final wordWidth =
            _estimateTextWidth(word, size, fontRef: segment.fontRef);
        final needsSpace = currentLine.isNotEmpty &&
            (currentLine.last.text.isNotEmpty &&
                !currentLine.last.text.endsWith(' '));
        final additionalWidth = needsSpace ? spaceWidth + wordWidth : wordWidth;

        if (currentLineWidth + additionalWidth > maxWidth &&
            currentLine.isNotEmpty) {
          // Start new line
          lines.add(currentLine);
          currentLine = <_TextSegment>[];
          currentLineWidth = 0.0;

          // Add word to new line with preserved formatting
          currentLine.add(_TextSegment(
            word,
            segment.fontRef,
            color: segment.color,
            underline: segment.underline,
            strikethrough: segment.strikethrough,
          ));
          currentLineWidth = wordWidth;
        } else {
          // Add to current line with preserved formatting
          if (needsSpace) {
            currentLine.add(_TextSegment(
              ' $word',
              segment.fontRef,
              color: segment.color,
              underline: segment.underline,
              strikethrough: segment.strikethrough,
            ));
            currentLineWidth += spaceWidth + wordWidth;
          } else {
            currentLine.add(_TextSegment(
              word,
              segment.fontRef,
              color: segment.color,
              underline: segment.underline,
              strikethrough: segment.strikethrough,
            ));
            currentLineWidth += wordWidth;
          }
        }
      }
    }

    // Add remaining line
    if (currentLine.isNotEmpty) {
      lines.add(currentLine);
    }

    // Ensure at least one empty line for empty paragraphs
    if (lines.isEmpty) {
      lines.add([_TextSegment('', '/F1')]);
    }

    return lines;
  }

  double _estimateTextWidth(String text, int size, {String? fontRef}) {
    // Courier (/F5) is monospaced: every glyph is 600/1000 em-wide.
    if (fontRef == '/F5') {
      return text.length * size * 0.6;
    }

    // For Helvetica (and its bold/italic variants) use the standard
    // Adobe glyph-width table (values in thousandths of an em-unit).
    // Bold variants are slightly wider on average but close enough for
    // layout; the same table is used for all Helvetica variants.
    double total = 0;
    for (final cu in text.codeUnits) {
      total += (_helveticaWidths[cu] ?? 500) / 1000 * size;
    }
    return total;
  }

  // Standard Adobe Helvetica glyph widths (WinAnsi / Latin-1 encoding).
  // Source: Adobe Font Metrics (AFM) for Helvetica.
  // Key: Unicode code-point, Value: width in thousandths of 1 em.
  static const _helveticaWidths = <int, int>{
    // ASCII printable 32-126
    0x20: 278, // space
    0x21: 278, // !
    0x22: 355, // "
    0x23: 556, // #
    0x24: 556, // $
    0x25: 889, // %
    0x26: 667, // &
    0x27: 191, // '
    0x28: 333, // (
    0x29: 333, // )
    0x2A: 389, // *
    0x2B: 584, // +
    0x2C: 278, // ,
    0x2D: 333, // -
    0x2E: 278, // .
    0x2F: 278, // /
    0x30: 556, // 0
    0x31: 556, // 1
    0x32: 556, // 2
    0x33: 556, // 3
    0x34: 556, // 4
    0x35: 556, // 5
    0x36: 556, // 6
    0x37: 556, // 7
    0x38: 556, // 8
    0x39: 556, // 9
    0x3A: 278, // :
    0x3B: 278, // ;
    0x3C: 584, // <
    0x3D: 584, // =
    0x3E: 584, // >
    0x3F: 556, // ?
    0x40: 1015, // @
    0x41: 667, // A
    0x42: 667, // B
    0x43: 722, // C
    0x44: 722, // D
    0x45: 667, // E
    0x46: 611, // F
    0x47: 778, // G
    0x48: 722, // H
    0x49: 278, // I
    0x4A: 500, // J
    0x4B: 667, // K
    0x4C: 556, // L
    0x4D: 833, // M
    0x4E: 722, // N
    0x4F: 778, // O
    0x50: 667, // P
    0x51: 778, // Q
    0x52: 722, // R
    0x53: 667, // S
    0x54: 611, // T
    0x55: 722, // U
    0x56: 667, // V
    0x57: 944, // W
    0x58: 667, // X
    0x59: 667, // Y
    0x5A: 611, // Z
    0x5B: 278, // [
    0x5C: 278, // backslash
    0x5D: 278, // ]
    0x5E: 469, // ^
    0x5F: 556, // _
    0x60: 333, // `
    0x61: 556, // a
    0x62: 556, // b
    0x63: 500, // c
    0x64: 556, // d
    0x65: 556, // e
    0x66: 278, // f
    0x67: 556, // g
    0x68: 556, // h
    0x69: 222, // i
    0x6A: 222, // j
    0x6B: 500, // k
    0x6C: 222, // l
    0x6D: 833, // m
    0x6E: 556, // n
    0x6F: 556, // o
    0x70: 556, // p
    0x71: 556, // q
    0x72: 333, // r
    0x73: 500, // s
    0x74: 278, // t
    0x75: 556, // u
    0x76: 500, // v
    0x77: 722, // w
    0x78: 500, // x
    0x79: 500, // y
    0x7A: 500, // z
    0x7B: 334, // {
    0x7C: 260, // |
    0x7D: 334, // }
    0x7E: 584, // ~
    // Latin-1 supplement (WinAnsi) – most common accented characters
    0xA0: 278, // non-breaking space
    0xA1: 333, // ¡
    0xA2: 556, // ¢
    0xA3: 556, // £
    0xA4: 556, // ¤
    0xA5: 556, // ¥
    0xA6: 260, // ¦
    0xA7: 556, // §
    0xA8: 333, // ¨
    0xA9: 737, // ©
    0xAA: 370, // ª
    0xAB: 556, // «
    0xAC: 584, // ¬
    0xAD: 333, // soft hyphen
    0xAE: 737, // ®
    0xAF: 333, // ¯
    0xB0: 400, // °
    0xB1: 584, // ±
    0xB2: 333, // ²
    0xB3: 333, // ³
    0xB4: 333, // ´
    0xB5: 556, // µ
    0xB6: 537, // ¶
    0xB7: 278, // ·
    0xB8: 333, // ¸
    0xB9: 333, // ¹
    0xBA: 365, // º
    0xBB: 556, // »
    0xBC: 834, // ¼
    0xBD: 834, // ½
    0xBE: 834, // ¾
    0xBF: 611, // ¿
    0xC0: 667, // À
    0xC1: 667, // Á
    0xC2: 667, // Â
    0xC3: 667, // Ã
    0xC4: 667, // Ä
    0xC5: 667, // Å
    0xC6: 1000, // Æ
    0xC7: 722, // Ç
    0xC8: 667, // È
    0xC9: 667, // É
    0xCA: 667, // Ê
    0xCB: 667, // Ë
    0xCC: 278, // Ì
    0xCD: 278, // Í
    0xCE: 278, // Î
    0xCF: 278, // Ï
    0xD0: 722, // Ð
    0xD1: 722, // Ñ
    0xD2: 778, // Ò
    0xD3: 778, // Ó
    0xD4: 778, // Ô
    0xD5: 778, // Õ
    0xD6: 778, // Ö
    0xD7: 584, // ×
    0xD8: 778, // Ø
    0xD9: 722, // Ù
    0xDA: 722, // Ú
    0xDB: 722, // Û
    0xDC: 722, // Ü
    0xDD: 667, // Ý
    0xDE: 667, // Þ
    0xDF: 611, // ß
    0xE0: 556, // à
    0xE1: 556, // á
    0xE2: 556, // â
    0xE3: 556, // ã
    0xE4: 556, // ä
    0xE5: 556, // å
    0xE6: 889, // æ
    0xE7: 500, // ç
    0xE8: 556, // è
    0xE9: 556, // é
    0xEA: 556, // ê
    0xEB: 556, // ë
    0xEC: 278, // ì
    0xED: 278, // í
    0xEE: 278, // î
    0xEF: 278, // ï
    0xF0: 556, // ð
    0xF1: 556, // ñ
    0xF2: 556, // ò
    0xF3: 556, // ó
    0xF4: 556, // ô
    0xF5: 556, // õ
    0xF6: 556, // ö
    0xF7: 584, // ÷
    0xF8: 611, // ø
    0xF9: 556, // ù
    0xFA: 556, // ú
    0xFB: 556, // û
    0xFC: 556, // ü
    0xFD: 500, // ý
    0xFE: 556, // þ
    0xFF: 500, // ÿ
    // Polish characters (mapped via WinAnsi)
    0x0104: 667, // Ą (like A)
    0x0105: 556, // ą (like a)
    0x0106: 722, // Ć (like C)
    0x0107: 500, // ć (like c)
    0x0118: 667, // Ę (like E)
    0x0119: 556, // ę (like e)
    0x0141: 556, // Ł (like L)
    0x0142: 222, // ł (like l)
    0x0143: 722, // Ń (like N)
    0x0144: 556, // ń (like n)
0x015A: 667, // Ś (like S)
    0x015B: 500, // ś (like s)
    0x0179: 611, // Ź (like Z)
    0x017A: 500, // ź (like z)
    0x017B: 611, // Ż (like Z)
    0x017C: 500, // ż (like z)
  };

  String _escapePdfString(String text) {
    final buffer = StringBuffer();
    for (final codeUnit in text.codeUnits) {
      switch (codeUnit) {
        // Basic escapes
        case 0x5C: // backslash
          buffer.write('\\\\');
        case 0x28: // (
          buffer.write('\\(');
        case 0x29: // )
          buffer.write('\\)');
        case 0x0A: // newline
          buffer.write('\\n');
        case 0x0D: // carriage return
          buffer.write('\\r');
        case 0x09: // tab
          buffer.write('\\t');

        // Typography (WinAnsi codes)
        case 0x2022: // bullet •
          buffer.write('\\225'); // WinAnsi 149
        case 0x2013: // en dash –
          buffer.write('\\226'); // WinAnsi 150
        case 0x2014: // em dash —
          buffer.write('\\227'); // WinAnsi 151
        case 0x2018: // left single quote '
          buffer.write('\\221'); // WinAnsi 145
        case 0x2019: // right single quote '
          buffer.write('\\222'); // WinAnsi 146
        case 0x201C: // left double quote "
          buffer.write('\\223'); // WinAnsi 147
        case 0x201D: // right double quote "
          buffer.write('\\224'); // WinAnsi 148
        case 0x2026: // ellipsis …
          buffer.write('\\205'); // WinAnsi 133
        case 0x20AC: // euro €
          buffer.write('\\200'); // WinAnsi 128
        case 0x2122: // trademark ™
          buffer.write('\\231'); // WinAnsi 153
        case 0x00A9: // copyright ©
          buffer.write('\\251'); // WinAnsi 169
        case 0x00AE: // registered ®
          buffer.write('\\256'); // WinAnsi 174

        // Polish characters (fallback to base letters - WinAnsi doesn't support these)
        case 0x0104: // Ą
          buffer.write('A');
        case 0x0105: // ą
          buffer.write('a');
        case 0x0106: // Ć
          buffer.write('C');
        case 0x0107: // ć
          buffer.write('c');
        case 0x0118: // Ę
          buffer.write('E');
        case 0x0119: // ę
          buffer.write('e');
        case 0x0141: // Ł
          buffer.write('L');
        case 0x0142: // ł
          buffer.write('l');
        case 0x0143: // Ń
          buffer.write('N');
        case 0x0144: // ń
          buffer.write('n');
        case 0x00D3: // Ó (in WinAnsi)
          buffer.write('\\323'); // WinAnsi 211
        case 0x00F3: // ó (in WinAnsi)
          buffer.write('\\363'); // WinAnsi 243
        case 0x015A: // Ś
          buffer.write('S');
        case 0x015B: // ś
          buffer.write('s');
        case 0x0179: // Ź
          buffer.write('Z');
        case 0x017A: // ź
          buffer.write('z');
        case 0x017B: // Ż
          buffer.write('Z');
        case 0x017C: // ż
          buffer.write('z');

        // German characters (in WinAnsi)
        case 0x00C4: // Ä
          buffer.write('\\304'); // WinAnsi 196
        case 0x00D6: // Ö
          buffer.write('\\326'); // WinAnsi 214
        case 0x00DC: // Ü
          buffer.write('\\334'); // WinAnsi 220
        case 0x00E4: // ä
          buffer.write('\\344'); // WinAnsi 228
        case 0x00F6: // ö
          buffer.write('\\366'); // WinAnsi 246
        case 0x00FC: // ü
          buffer.write('\\374'); // WinAnsi 252
        case 0x00DF: // ß
          buffer.write('\\337'); // WinAnsi 223

        // French characters (in WinAnsi)
        case 0x00C0: // À
          buffer.write('\\300'); // WinAnsi 192
        case 0x00C2: // Â
          buffer.write('\\302'); // WinAnsi 194
        case 0x00C7: // Ç
          buffer.write('\\307'); // WinAnsi 199
        case 0x00C8: // È
          buffer.write('\\310'); // WinAnsi 200
        case 0x00C9: // É
          buffer.write('\\311'); // WinAnsi 201
        case 0x00CA: // Ê
          buffer.write('\\312'); // WinAnsi 202
        case 0x00CB: // Ë
          buffer.write('\\313'); // WinAnsi 203
        case 0x00CE: // Î
          buffer.write('\\316'); // WinAnsi 206
        case 0x00CF: // Ï
          buffer.write('\\317'); // WinAnsi 207
        case 0x00D4: // Ô
          buffer.write('\\324'); // WinAnsi 212
        case 0x00D9: // Ù
          buffer.write('\\331'); // WinAnsi 217
        case 0x00DB: // Û
          buffer.write('\\333'); // WinAnsi 219
        case 0x00E0: // à
          buffer.write('\\340'); // WinAnsi 224
        case 0x00E2: // â
          buffer.write('\\342'); // WinAnsi 226
        case 0x00E7: // ç
          buffer.write('\\347'); // WinAnsi 231
        case 0x00E8: // è
          buffer.write('\\350'); // WinAnsi 232
        case 0x00E9: // é
          buffer.write('\\351'); // WinAnsi 233
        case 0x00EA: // ê
          buffer.write('\\352'); // WinAnsi 234
        case 0x00EB: // ë
          buffer.write('\\353'); // WinAnsi 235
        case 0x00EE: // î
          buffer.write('\\356'); // WinAnsi 238
        case 0x00EF: // ï
          buffer.write('\\357'); // WinAnsi 239
        case 0x00F4: // ô
          buffer.write('\\364'); // WinAnsi 244
        case 0x00F9: // ù
          buffer.write('\\371'); // WinAnsi 249
        case 0x00FB: // û
          buffer.write('\\373'); // WinAnsi 251

        default:
          if (codeUnit >= 32 && codeUnit <= 126) {
            // Standard ASCII printable
            buffer.writeCharCode(codeUnit);
          } else if (codeUnit >= 160 && codeUnit <= 255) {
            // Latin-1 supplement (maps to WinAnsi)
            buffer.write('\\${codeUnit.toRadixString(8).padLeft(3, '0')}');
          } else {
            // Unmappable - use replacement character or skip
            buffer.write('?');
          }
      }
    }
    return buffer.toString();
  }
}

/// Internal representation of a PDF object.
class _PdfObject {
  _PdfObject(this.id, this.content);

  final int id;
  final String content;
}

/// Internal representation of a text segment with font reference and formatting.
class _TextSegment {
  _TextSegment(
    this.text,
    this.fontRef, {
    this.color,
    this.underline = false,
    this.strikethrough = false,
  });

  final String text;
  final String fontRef;
  final String? color; // Hex color like "FF0000"
  final bool underline;
  final bool strikethrough;
}

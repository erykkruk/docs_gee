import '../models/models.dart';
import 'xml_utils.dart';

/// Generates document.xml (main content) for DOCX.
class DocumentXml {
  DocumentXml._();

  /// Counter for generating unique bookmark IDs.
  static int _bookmarkIdCounter = 0;

  /// Counter for generating unique hyperlink relationship IDs.
  static int _hyperlinkIdCounter = 0;

  /// Counter for generating unique image relationship IDs.
  static int _imageIdCounter = 0;

  /// First relationship ID reserved for embedded images.
  ///
  /// Kept clear of the fixed IDs (styles, numbering, header, footer) and of
  /// the hyperlink block that starts at 100.
  static const int _imageRelIdBase = 200;

  /// Generates the document.xml content.
  ///
  /// Returns the XML alongside the relationship maps the package parts need:
  /// hyperlink IDs to URLs, and image IDs to media part names.
  static ({
    String xml,
    Map<String, String> hyperlinks,
    Map<String, String> images,
  }) generate(
    DocxDocument document, {
    String? headerRelId,
    String? footerRelId,
  }) {
    final buffer = StringBuffer();
    final hyperlinks = <String, String>{};
    final images = <String, String>{};
    _bookmarkIdCounter = 0;
    _hyperlinkIdCounter = 0;
    _imageIdCounter = 0;

    buffer.writeln(XmlUtils.xmlDeclaration);
    buffer.writeln('<w:document xmlns:w="${XmlUtils.wNamespace}" '
        'xmlns:r="${XmlUtils.rNamespace}" '
        'xmlns:wp="${XmlUtils.wpNamespace}" '
        'xmlns:a="${XmlUtils.aNamespace}" '
        'xmlns:pic="${XmlUtils.picNamespace}">');
    buffer.writeln('  <w:body>');

    // Generate Table of Contents if enabled
    if (document.includeTableOfContents) {
      _writeTableOfContents(buffer, document.tocTitle, document.tocMaxLevel);
    }

    for (final item in document.content) {
      if (item is DocxParagraph) {
        writeParagraph(buffer, item, hyperlinks);
      } else if (item is DocxTable) {
        _writeTable(buffer, item, hyperlinks);
      } else if (item is DocxImage) {
        _writeImageParagraph(buffer, item, images);
      }
    }

    // Section properties (page size, margins). The header and footer
    // references have to come first inside sectPr: the schema fixes this
    // order and Word refuses the file otherwise.
    buffer.writeln('    <w:sectPr>');
    if (headerRelId != null) {
      buffer.writeln(
          '      <w:headerReference w:type="default" r:id="$headerRelId"/>');
    }
    if (footerRelId != null) {
      buffer.writeln(
          '      <w:footerReference w:type="default" r:id="$footerRelId"/>');
    }
    buffer.writeln('      <w:pgSz w:w="12240" w:h="15840"/>'); // Letter size
    buffer.writeln(
        '      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>');
    buffer.writeln('    </w:sectPr>');

    buffer.writeln('  </w:body>');
    buffer.writeln('</w:document>');
    return (xml: buffer.toString(), hyperlinks: hyperlinks, images: images);
  }

  /// Writes a Table of Contents field.
  static void _writeTableOfContents(
      StringBuffer buffer, String title, int maxLevel) {
    // TOC title
    buffer.writeln('    <w:p>');
    buffer.writeln('      <w:pPr>');
    buffer.writeln('        <w:pStyle w:val="Heading1"/>');
    buffer.writeln('      </w:pPr>');
    buffer.writeln('      <w:r>');
    buffer.writeln('        <w:t>${XmlUtils.escapeXml(title)}</w:t>');
    buffer.writeln('      </w:r>');
    buffer.writeln('    </w:p>');

    // TOC field
    buffer.writeln('    <w:p>');
    buffer.writeln('      <w:r>');
    buffer.writeln('        <w:fldChar w:fldCharType="begin"/>');
    buffer.writeln('      </w:r>');
    buffer.writeln('      <w:r>');
    buffer.writeln(
        '        <w:instrText xml:space="preserve"> TOC \\o "1-$maxLevel" \\h \\z \\u </w:instrText>');
    buffer.writeln('      </w:r>');
    buffer.writeln('      <w:r>');
    buffer.writeln('        <w:fldChar w:fldCharType="separate"/>');
    buffer.writeln('      </w:r>');
    buffer.writeln('      <w:r>');
    buffer.writeln('        <w:rPr><w:i/></w:rPr>');
    buffer.writeln(
        '        <w:t>Update this field to generate table of contents</w:t>');
    buffer.writeln('      </w:r>');
    buffer.writeln('      <w:r>');
    buffer.writeln('        <w:fldChar w:fldCharType="end"/>');
    buffer.writeln('      </w:r>');
    buffer.writeln('    </w:p>');

    // Page break after TOC
    buffer.writeln('    <w:p>');
    buffer.writeln('      <w:r>');
    buffer.writeln('        <w:br w:type="page"/>');
    buffer.writeln('      </w:r>');
    buffer.writeln('    </w:p>');
  }

  /// Writes a body-level paragraph.
  ///
  /// Library-internal (`xml_parts` is not exported): also used by
  /// [HeaderFooterXml], which passes `allowHyperlinks: false` because header
  /// and footer parts carry their own relationship file.
  static void writeParagraph(
    StringBuffer buffer,
    DocxParagraph paragraph,
    Map<String, String> hyperlinks, {
    bool allowHyperlinks = true,
  }) {
    // Insert a separate page break paragraph before this one if requested
    if (paragraph.pageBreakBefore) {
      buffer.writeln('    <w:p>');
      buffer.writeln('      <w:r>');
      buffer.writeln('        <w:br w:type="page"/>');
      buffer.writeln('      </w:r>');
      buffer.writeln('    </w:p>');
    }

    buffer.writeln('    <w:p>');

    // Write bookmark start if paragraph has a bookmark
    String? bookmarkId;
    if (paragraph.bookmarkName != null) {
      bookmarkId = _bookmarkIdCounter.toString();
      _bookmarkIdCounter++;
      buffer.writeln(
          '      <w:bookmarkStart w:id="$bookmarkId" w:name="${XmlUtils.escapeXml(paragraph.bookmarkName!)}"/>');
    }

    // Paragraph properties
    final hasStyle = paragraph.style != DocxParagraphStyle.normal;
    final hasAlignment = paragraph.alignment != DocxAlignment.left;
    final hasIndent = paragraph.indentLevel > 0 && paragraph.style.isList;
    final hasParagraphProps =
        hasStyle || hasAlignment || hasIndent || paragraph.rtl;

    if (hasParagraphProps) {
      buffer.writeln('      <w:pPr>');

      if (hasStyle) {
        buffer.writeln(
          '        <w:pStyle w:val="${paragraph.style.styleId}"/>',
        );
      }

      // Add numPr with ilvl for nested lists
      if (hasIndent) {
        final numId = _getNumIdForStyle(paragraph.style);
        buffer.writeln('        <w:numPr>');
        buffer.writeln('          <w:ilvl w:val="${paragraph.indentLevel}"/>');
        buffer.writeln('          <w:numId w:val="$numId"/>');
        buffer.writeln('        </w:numPr>');
      }

      // Right-to-left paragraph direction. Written before <w:jc> so the
      // alignment value is interpreted in the flipped direction.
      if (paragraph.rtl) {
        buffer.writeln('        <w:bidi/>');
      }

      if (hasAlignment) {
        buffer.writeln(
          '        <w:jc w:val="${paragraph.alignment.value}"/>',
        );
      }

      buffer.writeln('      </w:pPr>');
    }

    // Write runs
    for (final run in paragraph.runs) {
      _writeRun(buffer, run, hyperlinks, allowHyperlinks: allowHyperlinks);
    }

    // Write bookmark end if paragraph has a bookmark
    if (bookmarkId != null) {
      buffer.writeln('      <w:bookmarkEnd w:id="$bookmarkId"/>');
    }

    buffer.writeln('    </w:p>');
  }

  /// Returns the numId for a given list style.
  static int _getNumIdForStyle(DocxParagraphStyle style) {
    return switch (style) {
      DocxParagraphStyle.listBullet => 1,
      DocxParagraphStyle.listNumber => 2,
      DocxParagraphStyle.listDash => 3,
      DocxParagraphStyle.listNumberAlpha => 4,
      DocxParagraphStyle.listNumberRoman => 5,
      _ => 1,
    };
  }

  static void _writeRun(
    StringBuffer buffer,
    DocxRun run,
    Map<String, String> hyperlinks, {
    bool allowHyperlinks = true,
  }) {
    // Handle line break runs
    if (run.isLineBreak) {
      buffer.writeln('      <w:r><w:br/></w:r>');
      return;
    }

    // Handle field runs (page number, page count)
    if (run.isField) {
      buffer.write('      ');
      _writeFieldRun(buffer, run);
      buffer.writeln();
      return;
    }

    // Handle external hyperlinks
    if (run.hyperlink != null && allowHyperlinks) {
      final rId = 'rId${100 + _hyperlinkIdCounter}';
      _hyperlinkIdCounter++;
      hyperlinks[rId] = run.hyperlink!;
      buffer.write('      <w:hyperlink r:id="$rId">');
      _writeRunContent(buffer, run, isHyperlink: true);
      buffer.writeln('</w:hyperlink>');
      return;
    }

    // Handle internal bookmark references
    if (run.bookmarkRef != null) {
      buffer.write(
          '      <w:hyperlink w:anchor="${XmlUtils.escapeXml(run.bookmarkRef!)}">');
      _writeRunContent(buffer, run, isHyperlink: true);
      buffer.writeln('</w:hyperlink>');
      return;
    }

    // Regular run
    buffer.write('      ');
    _writeRunContent(buffer, run, isHyperlink: false);
    buffer.writeln();
  }

  /// Writes the run content (used by both regular runs and hyperlinks).
  static void _writeRunContent(StringBuffer buffer, DocxRun run,
      {required bool isHyperlink}) {
    buffer.write('<w:r>');

    // Run properties (formatting)
    final hasLinkStyle = isHyperlink && !run.underline && run.color == null;
    if (run.hasFormatting || hasLinkStyle) {
      buffer.write('<w:rPr>');
      if (run.bold) buffer.write('<w:b/>');
      if (run.italic) buffer.write('<w:i/>');
      if (run.underline || hasLinkStyle) {
        buffer.write('<w:u w:val="single"/>');
      }
      if (run.strikethrough) buffer.write('<w:strike/>');
      if (run.script != DocxScript.baseline) {
        buffer.write('<w:vertAlign w:val="${run.script.value}"/>');
      }
      if (run.fontSize != null) {
        // DOCX stores font size in half-points, so 14pt is w:val="28".
        final halfPoints = run.fontSize! * 2;
        buffer.write('<w:sz w:val="$halfPoints"/>');
        buffer.write('<w:szCs w:val="$halfPoints"/>');
      }
      if (run.rtl) {
        buffer.write('<w:rtl/>');
      }
      if (run.color != null) {
        buffer.write('<w:color w:val="${run.color}"/>');
      } else if (hasLinkStyle) {
        buffer.write('<w:color w:val="0000FF"/>'); // Blue for links
      }
      if (run.backgroundColor != null) {
        buffer.write(
            '<w:highlight w:val="${_mapHighlightColor(run.backgroundColor!)}"/>');
      }
      buffer.write('</w:rPr>');
    }

    // Text content - handle line breaks (\n) by splitting into multiple <w:t> with <w:br/>
    final lines = run.text.split('\n');
    for (int i = 0; i < lines.length; i++) {
      final line = lines[i];
      final escapedText = XmlUtils.escapeXml(line);
      // Use xml:space="preserve" if text has leading/trailing spaces
      if (line.startsWith(' ') || line.endsWith(' ')) {
        buffer.write('<w:t xml:space="preserve">$escapedText</w:t>');
      } else {
        buffer.write('<w:t>$escapedText</w:t>');
      }
      // Add line break between lines (not after the last one)
      if (i < lines.length - 1) {
        buffer.write('<w:br/>');
      }
    }

    buffer.write('</w:r>');
  }

  /// Maps hex color to Word highlight color name.
  /// Word only supports a limited set of highlight colors.
  static String _mapHighlightColor(String hexColor) {
    final hex = hexColor.toUpperCase();
    return switch (hex) {
      'FFFF00' || 'FFD700' => 'yellow',
      '00FF00' || '90EE90' => 'green',
      '00FFFF' || 'E0FFFF' => 'cyan',
      'FF00FF' || 'FF69B4' => 'magenta',
      '0000FF' || '4169E1' => 'blue',
      'FF0000' || 'DC143C' => 'red',
      '000080' => 'darkBlue',
      '008080' => 'darkCyan',
      '008000' => 'darkGreen',
      '800080' => 'darkMagenta',
      '800000' => 'darkRed',
      '808000' => 'darkYellow',
      '808080' => 'darkGray',
      'C0C0C0' || 'D3D3D3' => 'lightGray',
      '000000' => 'black',
      'FFFFFF' => 'white',
      _ => 'yellow', // Default to yellow for unknown colors
    };
  }

  // ============================================
  // TABLE GENERATION
  // ============================================

  /// Total usable table width in twips (6.5 inches at 1440 twips/inch).
  static const int _tableWidth = 9360;

  /// Calculates column widths in twips based on table configuration.
  static List<int> _calculateColumnWidths(DocxTable table) {
    final colCount = table.columnCount;
    if (colCount == 0) return [];

    if (table.columnWidths != null && table.columnWidths!.length == colCount) {
      // Use custom widths (percentages converted to twips)
      return table.columnWidths!
          .map((percent) => (_tableWidth * percent / 100).round())
          .toList();
    } else {
      // Even distribution
      final width = _tableWidth ~/ colCount;
      return List.filled(colCount, width);
    }
  }

  /// Writes a table element to the buffer.
  static void _writeTable(
      StringBuffer buffer, DocxTable table, Map<String, String> hyperlinks) {
    buffer.writeln('    <w:tbl>');

    // Table properties
    buffer.writeln('      <w:tblPr>');
    buffer.writeln('        <w:tblW w:w="0" w:type="auto"/>');

    // Borders
    if (table.borders.hasBorders) {
      buffer.writeln('        <w:tblBorders>');
      _writeTableBorder(buffer, 'top', table.borders.top);
      _writeTableBorder(buffer, 'left', table.borders.left);
      _writeTableBorder(buffer, 'bottom', table.borders.bottom);
      _writeTableBorder(buffer, 'right', table.borders.right);
      _writeTableBorder(buffer, 'insideH', table.borders.insideH);
      _writeTableBorder(buffer, 'insideV', table.borders.insideV);
      buffer.writeln('        </w:tblBorders>');
    }

    // Table-wide cell margins, applied to every cell without its own padding.
    final cellPadding = table.cellPadding;
    if (cellPadding != null) {
      buffer.writeln('        <w:tblCellMar>');
      _writeCellMarginEdges(buffer, cellPadding, '          ');
      buffer.writeln('        </w:tblCellMar>');
    }

    buffer.writeln('      </w:tblPr>');

    // Calculate column widths
    final columnWidths = _calculateColumnWidths(table);

    // Table grid (column definitions)
    if (columnWidths.isNotEmpty) {
      buffer.writeln('      <w:tblGrid>');
      for (final width in columnWidths) {
        buffer.writeln('        <w:gridCol w:w="$width"/>');
      }
      buffer.writeln('      </w:tblGrid>');
    }

    // Rows
    for (final row in table.rows) {
      _writeTableRow(buffer, row, columnWidths, hyperlinks);
    }

    buffer.writeln('    </w:tbl>');
  }

  /// Writes a table border element.
  static void _writeTableBorder(
      StringBuffer buffer, String name, DocxBorder? border) {
    if (border == null) {
      buffer.writeln('          <w:$name w:val="nil"/>');
    } else {
      buffer.writeln(
          '          <w:$name w:val="${border.style.value}" w:sz="${border.size}" w:space="0" w:color="${border.color}"/>');
    }
  }

  /// Writes a table row element.
  static void _writeTableRow(StringBuffer buffer, DocxTableRow row,
      List<int> columnWidths, Map<String, String> hyperlinks) {
    buffer.writeln('      <w:tr>');

    int colIndex = 0;
    for (final cell in row.cells) {
      // Calculate cell width (sum of spanned columns)
      int cellWidth = 0;
      for (int i = 0;
          i < cell.colSpan && colIndex + i < columnWidths.length;
          i++) {
        cellWidth += columnWidths[colIndex + i];
      }
      if (cellWidth == 0 && columnWidths.isNotEmpty) {
        cellWidth = columnWidths.first;
      }

      _writeTableCell(buffer, cell, cellWidth, hyperlinks);
      colIndex += cell.colSpan;
    }

    buffer.writeln('      </w:tr>');
  }

  /// Writes a table cell element.
  static void _writeTableCell(StringBuffer buffer, DocxTableCell cell,
      int width, Map<String, String> hyperlinks) {
    buffer.writeln('        <w:tc>');

    // Cell properties
    buffer.writeln('          <w:tcPr>');
    buffer.writeln('            <w:tcW w:w="$width" w:type="dxa"/>');

    // Horizontal merge (colspan)
    if (cell.colSpan > 1) {
      buffer.writeln('            <w:gridSpan w:val="${cell.colSpan}"/>');
    }

    // Vertical merge (rowspan)
    if (cell.rowSpan > 1) {
      buffer.writeln('            <w:vMerge w:val="restart"/>');
    } else if (cell.isMergedContinuation) {
      buffer.writeln('            <w:vMerge/>');
    }

    // Cell borders (override table-level borders)
    if (cell.borders != null && cell.borders!.hasBorders) {
      buffer.writeln('            <w:tcBorders>');
      _writeTableBorder(buffer, 'top', cell.borders!.top);
      _writeTableBorder(buffer, 'left', cell.borders!.left);
      _writeTableBorder(buffer, 'bottom', cell.borders!.bottom);
      _writeTableBorder(buffer, 'right', cell.borders!.right);
      buffer.writeln('            </w:tcBorders>');
    }

    // Background color
    if (cell.backgroundColor != null) {
      buffer.writeln(
          '            <w:shd w:val="clear" w:fill="${cell.backgroundColor}"/>');
    }

    // Vertical alignment
    if (cell.verticalAlignment != DocxVerticalAlignment.top) {
      buffer.writeln(
          '            <w:vAlign w:val="${cell.verticalAlignment.value}"/>');
    }

    // Per-cell padding, overriding the table-wide margins.
    final padding = cell.padding;
    if (padding != null) {
      buffer.writeln('            <w:tcMar>');
      _writeCellMarginEdges(buffer, padding, '              ');
      buffer.writeln('            </w:tcMar>');
    }

    buffer.writeln('          </w:tcPr>');

    // Cell content (paragraphs)
    if (cell.paragraphs.isEmpty) {
      // Empty cell needs at least one empty paragraph
      buffer.writeln('          <w:p/>');
    } else {
      for (final paragraph in cell.paragraphs) {
        _writeCellParagraph(buffer, paragraph, hyperlinks);
      }
    }

    buffer.writeln('        </w:tc>');
  }

  /// Writes a paragraph inside a table cell (with adjusted indentation).
  static void _writeCellParagraph(StringBuffer buffer, DocxParagraph paragraph,
      Map<String, String> hyperlinks) {
    // Insert a separate page break paragraph before this one if requested
    if (paragraph.pageBreakBefore) {
      buffer.writeln('          <w:p>');
      buffer.writeln('            <w:r>');
      buffer.writeln('              <w:br w:type="page"/>');
      buffer.writeln('            </w:r>');
      buffer.writeln('          </w:p>');
    }

    buffer.writeln('          <w:p>');

    // Write bookmark start if paragraph has a bookmark
    String? bookmarkId;
    if (paragraph.bookmarkName != null) {
      bookmarkId = _bookmarkIdCounter.toString();
      _bookmarkIdCounter++;
      buffer.writeln(
          '            <w:bookmarkStart w:id="$bookmarkId" w:name="${XmlUtils.escapeXml(paragraph.bookmarkName!)}"/>');
    }

    // Paragraph properties
    final hasStyle = paragraph.style != DocxParagraphStyle.normal;
    final hasAlignment = paragraph.alignment != DocxAlignment.left;
    final hasIndent = paragraph.indentLevel > 0 && paragraph.style.isList;
    final hasParagraphProps =
        hasStyle || hasAlignment || hasIndent || paragraph.rtl;

    if (hasParagraphProps) {
      buffer.writeln('            <w:pPr>');

      if (hasStyle) {
        buffer.writeln(
          '              <w:pStyle w:val="${paragraph.style.styleId}"/>',
        );
      }

      if (hasIndent) {
        final numId = _getNumIdForStyle(paragraph.style);
        buffer.writeln('              <w:numPr>');
        buffer.writeln(
            '                <w:ilvl w:val="${paragraph.indentLevel}"/>');
        buffer.writeln('                <w:numId w:val="$numId"/>');
        buffer.writeln('              </w:numPr>');
      }

      if (paragraph.rtl) {
        buffer.writeln('              <w:bidi/>');
      }

      if (hasAlignment) {
        buffer.writeln(
          '              <w:jc w:val="${paragraph.alignment.value}"/>',
        );
      }

      buffer.writeln('            </w:pPr>');
    }

    // Write runs
    for (final run in paragraph.runs) {
      _writeCellRun(buffer, run, hyperlinks);
    }

    // Write bookmark end if paragraph has a bookmark
    if (bookmarkId != null) {
      buffer.writeln('            <w:bookmarkEnd w:id="$bookmarkId"/>');
    }

    buffer.writeln('          </w:p>');
  }

  /// Writes a run inside a table cell (with adjusted indentation).
  static void _writeCellRun(
      StringBuffer buffer, DocxRun run, Map<String, String> hyperlinks) {
    // Handle line break runs
    if (run.isLineBreak) {
      buffer.writeln('            <w:r><w:br/></w:r>');
      return;
    }

    // Handle field runs (page number, page count)
    if (run.isField) {
      buffer.write('            ');
      _writeFieldRun(buffer, run);
      buffer.writeln();
      return;
    }

    // Handle external hyperlinks
    if (run.hyperlink != null) {
      final rId = 'rId${100 + _hyperlinkIdCounter}';
      _hyperlinkIdCounter++;
      hyperlinks[rId] = run.hyperlink!;
      buffer.write('            <w:hyperlink r:id="$rId">');
      _writeRunContent(buffer, run, isHyperlink: true);
      buffer.writeln('</w:hyperlink>');
      return;
    }

    // Handle internal bookmark references
    if (run.bookmarkRef != null) {
      buffer.write(
          '            <w:hyperlink w:anchor="${XmlUtils.escapeXml(run.bookmarkRef!)}">');
      _writeRunContent(buffer, run, isHyperlink: true);
      buffer.writeln('</w:hyperlink>');
      return;
    }

    // Regular run
    buffer.write('            ');
    _writeRunContent(buffer, run, isHyperlink: false);
    buffer.writeln();
  }

  /// Writes the four `<w:top>`/`<w:left>`/`<w:bottom>`/`<w:right>` margin
  /// edges shared by `<w:tblCellMar>` and `<w:tcMar>`.
  static void _writeCellMarginEdges(
    StringBuffer buffer,
    DocxCellPadding padding,
    String indent,
  ) {
    buffer.writeln('$indent<w:top w:w="${padding.top}" w:type="dxa"/>');
    buffer.writeln('$indent<w:left w:w="${padding.left}" w:type="dxa"/>');
    buffer.writeln('$indent<w:bottom w:w="${padding.bottom}" w:type="dxa"/>');
    buffer.writeln('$indent<w:right w:w="${padding.right}" w:type="dxa"/>');
  }

  /// Writes a field run as `<w:fldSimple>`.
  ///
  /// The nested run holds a placeholder value so the field still shows
  /// something in readers that do not recalculate fields; Word replaces it
  /// on open.
  static void _writeFieldRun(StringBuffer buffer, DocxRun run) {
    final field = run.field!;
    buffer.write('<w:fldSimple w:instr=" ${field.instruction} ">');
    buffer.write('<w:r>');
    final hasProps =
        run.bold || run.italic || run.color != null || run.fontSize != null;
    if (hasProps) {
      buffer.write('<w:rPr>');
      if (run.bold) buffer.write('<w:b/>');
      if (run.italic) buffer.write('<w:i/>');
      if (run.color != null) {
        buffer.write('<w:color w:val="${run.color}"/>');
      }
      if (run.fontSize != null) {
        final halfPoints = run.fontSize! * 2;
        buffer.write('<w:sz w:val="$halfPoints"/>');
        buffer.write('<w:szCs w:val="$halfPoints"/>');
      }
      buffer.write('</w:rPr>');
    }
    buffer.write('<w:t>1</w:t>');
    buffer.write('</w:r>');
    buffer.write('</w:fldSimple>');
  }

  // ============================================
  // IMAGE GENERATION
  // ============================================

  /// Writes an image as its own paragraph, registering the media part in
  /// [images] so the relationship file and the archive can pick it up.
  static void _writeImageParagraph(
    StringBuffer buffer,
    DocxImage image,
    Map<String, String> images,
  ) {
    final index = _imageIdCounter;
    _imageIdCounter++;
    final rId = 'rId${_imageRelIdBase + index}';
    final fileName = 'image${index + 1}.${image.format.extension}';
    images[rId] = 'media/$fileName';

    buffer.writeln('    <w:p>');
    if (image.alignment != DocxAlignment.left) {
      buffer.writeln('      <w:pPr>');
      buffer.writeln('        <w:jc w:val="${image.alignment.value}"/>');
      buffer.writeln('      </w:pPr>');
    }
    buffer.writeln('      <w:r>');
    _writeDrawing(buffer, image, rId: rId, id: index + 1, fileName: fileName);
    buffer.writeln('      </w:r>');
    buffer.writeln('    </w:p>');
  }

  /// Writes the inline DrawingML markup for one picture.
  static void _writeDrawing(
    StringBuffer buffer,
    DocxImage image, {
    required String rId,
    required int id,
    required String fileName,
  }) {
    final cx = image.widthEmu;
    final cy = image.heightEmu;
    final name = XmlUtils.escapeXml(fileName);
    final descr = image.altText == null
        ? ''
        : ' descr="${XmlUtils.escapeXml(image.altText!)}"';

    buffer.writeln('        <w:drawing>');
    buffer.writeln(
        '          <wp:inline distT="0" distB="0" distL="0" distR="0">');
    buffer.writeln('            <wp:extent cx="$cx" cy="$cy"/>');
    buffer.writeln('            <wp:effectExtent l="0" t="0" r="0" b="0"/>');
    buffer.writeln('            <wp:docPr id="$id" name="Picture $id"$descr/>');
    buffer.writeln('            <wp:cNvGraphicFramePr>');
    buffer.writeln('              <a:graphicFrameLocks noChangeAspect="1"/>');
    buffer.writeln('            </wp:cNvGraphicFramePr>');
    buffer.writeln('            <a:graphic>');
    buffer.writeln(
        '              <a:graphicData uri="${XmlUtils.picNamespace}">');
    buffer.writeln('                <pic:pic>');
    buffer.writeln('                  <pic:nvPicPr>');
    buffer.writeln(
        '                    <pic:cNvPr id="$id" name="$name"$descr/>');
    buffer.writeln('                    <pic:cNvPicPr/>');
    buffer.writeln('                  </pic:nvPicPr>');
    buffer.writeln('                  <pic:blipFill>');
    buffer.writeln('                    <a:blip r:embed="$rId"/>');
    buffer.writeln('                    <a:stretch>');
    buffer.writeln('                      <a:fillRect/>');
    buffer.writeln('                    </a:stretch>');
    buffer.writeln('                  </pic:blipFill>');
    buffer.writeln('                  <pic:spPr>');
    buffer.writeln('                    <a:xfrm>');
    buffer.writeln('                      <a:off x="0" y="0"/>');
    buffer.writeln('                      <a:ext cx="$cx" cy="$cy"/>');
    buffer.writeln('                    </a:xfrm>');
    buffer.writeln('                    <a:prstGeom prst="rect">');
    buffer.writeln('                      <a:avLst/>');
    buffer.writeln('                    </a:prstGeom>');
    buffer.writeln('                  </pic:spPr>');
    buffer.writeln('                </pic:pic>');
    buffer.writeln('              </a:graphicData>');
    buffer.writeln('            </a:graphic>');
    buffer.writeln('          </wp:inline>');
    buffer.writeln('        </w:drawing>');
  }
}

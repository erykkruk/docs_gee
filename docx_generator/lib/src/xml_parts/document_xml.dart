import '../models/models.dart';
import 'xml_utils.dart';

/// Generates document.xml (main content) for DOCX.
class DocumentXml {
  DocumentXml._();

  /// Generates the document.xml content.
  static String generate(DocxDocument document) {
    final buffer = StringBuffer();
    buffer.writeln(XmlUtils.xmlDeclaration);
    buffer.writeln('<w:document xmlns:w="${XmlUtils.wNamespace}">');
    buffer.writeln('  <w:body>');

    for (final item in document.content) {
      if (item is DocxParagraph) {
        _writeParagraph(buffer, item);
      } else if (item is DocxTable) {
        _writeTable(buffer, item);
      }
    }

    // Section properties (page size, margins)
    buffer.writeln('    <w:sectPr>');
    buffer.writeln('      <w:pgSz w:w="12240" w:h="15840"/>'); // Letter size
    buffer.writeln(
        '      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>');
    buffer.writeln('    </w:sectPr>');

    buffer.writeln('  </w:body>');
    buffer.writeln('</w:document>');
    return buffer.toString();
  }

  static void _writeParagraph(StringBuffer buffer, DocxParagraph paragraph) {
    buffer.writeln('    <w:p>');

    // Paragraph properties
    final hasStyle = paragraph.style != DocxParagraphStyle.normal;
    final hasAlignment = paragraph.alignment != DocxAlignment.left;
    final hasIndent = paragraph.indentLevel > 0 && paragraph.style.isList;
    final hasParagraphProps =
        hasStyle || hasAlignment || paragraph.pageBreakBefore || hasIndent;

    if (hasParagraphProps) {
      buffer.writeln('      <w:pPr>');

      if (paragraph.pageBreakBefore) {
        buffer.writeln('        <w:pageBreakBefore/>');
      }

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

      if (hasAlignment) {
        buffer.writeln(
          '        <w:jc w:val="${paragraph.alignment.value}"/>',
        );
      }

      buffer.writeln('      </w:pPr>');
    }

    // Write runs
    for (final run in paragraph.runs) {
      _writeRun(buffer, run);
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

  static void _writeRun(StringBuffer buffer, DocxRun run) {
    buffer.write('      <w:r>');

    // Run properties (formatting)
    if (run.hasFormatting) {
      buffer.write('<w:rPr>');
      if (run.bold) buffer.write('<w:b/>');
      if (run.italic) buffer.write('<w:i/>');
      if (run.underline) buffer.write('<w:u w:val="single"/>');
      if (run.strikethrough) buffer.write('<w:strike/>');
      if (run.color != null) {
        buffer.write('<w:color w:val="${run.color}"/>');
      }
      if (run.backgroundColor != null) {
        buffer.write(
            '<w:highlight w:val="${_mapHighlightColor(run.backgroundColor!)}"/>');
        // Alternative: shading for exact colors
        // buffer.write('<w:shd w:val="clear" w:fill="${run.backgroundColor}"/>');
      }
      buffer.write('</w:rPr>');
    }

    // Text content
    final escapedText = XmlUtils.escapeXml(run.text);
    // Use xml:space="preserve" if text has leading/trailing spaces
    if (run.text.startsWith(' ') || run.text.endsWith(' ')) {
      buffer.write('<w:t xml:space="preserve">$escapedText</w:t>');
    } else {
      buffer.write('<w:t>$escapedText</w:t>');
    }

    buffer.writeln('</w:r>');
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

  /// Writes a table element to the buffer.
  static void _writeTable(StringBuffer buffer, DocxTable table) {
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

    buffer.writeln('      </w:tblPr>');

    // Table grid (column definitions)
    if (table.columnCount > 0) {
      buffer.writeln('      <w:tblGrid>');
      // Total usable width ~9360 twips (6.5 inches at 1440 twips/inch)
      final colWidth = 9360 ~/ table.columnCount;
      for (int i = 0; i < table.columnCount; i++) {
        buffer.writeln('        <w:gridCol w:w="$colWidth"/>');
      }
      buffer.writeln('      </w:tblGrid>');
    }

    // Rows
    for (final row in table.rows) {
      _writeTableRow(buffer, row, table.columnCount);
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
  static void _writeTableRow(
      StringBuffer buffer, DocxTableRow row, int columnCount) {
    buffer.writeln('      <w:tr>');

    final colWidth = columnCount > 0 ? 9360 ~/ columnCount : 9360;
    for (final cell in row.cells) {
      _writeTableCell(buffer, cell, colWidth);
    }

    buffer.writeln('      </w:tr>');
  }

  /// Writes a table cell element.
  static void _writeTableCell(
      StringBuffer buffer, DocxTableCell cell, int width) {
    buffer.writeln('        <w:tc>');

    // Cell properties
    buffer.writeln('          <w:tcPr>');
    buffer.writeln('            <w:tcW w:w="$width" w:type="dxa"/>');

    if (cell.backgroundColor != null) {
      buffer.writeln(
          '            <w:shd w:val="clear" w:fill="${cell.backgroundColor}"/>');
    }

    buffer.writeln('          </w:tcPr>');

    // Cell content (paragraphs)
    if (cell.paragraphs.isEmpty) {
      // Empty cell needs at least one empty paragraph
      buffer.writeln('          <w:p/>');
    } else {
      for (final paragraph in cell.paragraphs) {
        _writeCellParagraph(buffer, paragraph);
      }
    }

    buffer.writeln('        </w:tc>');
  }

  /// Writes a paragraph inside a table cell (with adjusted indentation).
  static void _writeCellParagraph(StringBuffer buffer, DocxParagraph paragraph) {
    buffer.writeln('          <w:p>');

    // Paragraph properties
    final hasStyle = paragraph.style != DocxParagraphStyle.normal;
    final hasAlignment = paragraph.alignment != DocxAlignment.left;
    final hasIndent = paragraph.indentLevel > 0 && paragraph.style.isList;
    final hasParagraphProps =
        hasStyle || hasAlignment || paragraph.pageBreakBefore || hasIndent;

    if (hasParagraphProps) {
      buffer.writeln('            <w:pPr>');

      if (paragraph.pageBreakBefore) {
        buffer.writeln('              <w:pageBreakBefore/>');
      }

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

      if (hasAlignment) {
        buffer.writeln(
          '              <w:jc w:val="${paragraph.alignment.value}"/>',
        );
      }

      buffer.writeln('            </w:pPr>');
    }

    // Write runs
    for (final run in paragraph.runs) {
      _writeCellRun(buffer, run);
    }

    buffer.writeln('          </w:p>');
  }

  /// Writes a run inside a table cell (with adjusted indentation).
  static void _writeCellRun(StringBuffer buffer, DocxRun run) {
    buffer.write('            <w:r>');

    // Run properties (formatting)
    if (run.hasFormatting) {
      buffer.write('<w:rPr>');
      if (run.bold) buffer.write('<w:b/>');
      if (run.italic) buffer.write('<w:i/>');
      if (run.underline) buffer.write('<w:u w:val="single"/>');
      if (run.strikethrough) buffer.write('<w:strike/>');
      if (run.color != null) {
        buffer.write('<w:color w:val="${run.color}"/>');
      }
      if (run.backgroundColor != null) {
        buffer.write(
            '<w:highlight w:val="${_mapHighlightColor(run.backgroundColor!)}"/>');
      }
      buffer.write('</w:rPr>');
    }

    // Text content
    final escapedText = XmlUtils.escapeXml(run.text);
    if (run.text.startsWith(' ') || run.text.endsWith(' ')) {
      buffer.write('<w:t xml:space="preserve">$escapedText</w:t>');
    } else {
      buffer.write('<w:t>$escapedText</w:t>');
    }

    buffer.writeln('</w:r>');
  }
}

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

    // Object 7: Info dictionary (metadata)
    final infoObj = _PdfObject(7, _buildInfo(document));
    _objects.add(infoObj);

    // Create page objects and content streams
    int nextObjId = 8;
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
    _writeBytes(buffer, '/Info 7 0 R\n');
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
    // Page objects start at ID 8, every other object (page, content, page, content...)
    for (int i = 0; i < pageCount; i++) {
      final pageObjId = 8 + (i * 2);
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
        '  >>\n'
        '>>\n'
        '>>';
  }

  String _buildStreamObject(String content) {
    final bytes = utf8.encode(content);
    return '<<\n/Length ${bytes.length}\n>>\nstream\n$content\nendstream';
  }

  List<List<DocxParagraph>> _paginateDocument(DocxDocument document) {
    final pages = <List<DocxParagraph>>[];
    var currentPage = <DocxParagraph>[];
    var currentY = _contentHeight;

    for (final paragraph in document.paragraphs) {
      // Check for explicit page break
      if (paragraph.pageBreakBefore && currentPage.isNotEmpty) {
        pages.add(currentPage);
        currentPage = <DocxParagraph>[];
        currentY = _contentHeight;
      }

      final paragraphHeight = _estimateParagraphHeight(paragraph);

      // Check if paragraph fits on current page
      if (currentY - paragraphHeight < 0 && currentPage.isNotEmpty) {
        pages.add(currentPage);
        currentPage = <DocxParagraph>[];
        currentY = _contentHeight;
      }

      currentPage.add(paragraph);
      currentY -= paragraphHeight;
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
      DocxParagraphStyle.quote => fontSize,
      _ => fontSize,
    };
  }

  String _buildContentStream(List<DocxParagraph> paragraphs) {
    final buffer = StringBuffer();

    buffer.writeln('BT'); // Begin text

    var currentY = pageHeight - marginTop;

    // Track counters for each list type and level
    // Key: "styleType_level", Value: current counter
    final listCounters = <String, int>{};
    DocxParagraphStyle? lastListStyle;
    int lastIndentLevel = -1;

    for (final paragraph in paragraphs) {
      final size = _getFontSizeForStyle(paragraph.style);
      final lineHeight = size * 1.4;

      // Handle list prefixes and indentation
      String prefix = '';
      // Calculate indent offset in points (18 points per level, ~0.25 inch)
      final indentOffset = paragraph.indentLevel * 18.0;

      if (paragraph.style.isList) {
        // Reset counters if switching list type at same level
        if (lastListStyle != paragraph.style ||
            paragraph.indentLevel < lastIndentLevel) {
          // Reset all counters at this level and deeper
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
          prefix = '  \x95 '; // Bullet character
        } else if (paragraph.style == DocxParagraphStyle.listDash) {
          prefix = '  - '; // Dash character
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

        lastListStyle = paragraph.style;
        lastIndentLevel = paragraph.indentLevel;
      } else {
        // Reset all list counters when not in a list
        listCounters.clear();
        lastListStyle = null;
        lastIndentLevel = -1;
      }

      // Combine all runs into text segments with formatting info
      final segments = <_TextSegment>[];

      // Determine if heading (bold by default)
      final isHeading = paragraph.style == DocxParagraphStyle.heading1 ||
          paragraph.style == DocxParagraphStyle.heading2 ||
          paragraph.style == DocxParagraphStyle.heading3 ||
          paragraph.style == DocxParagraphStyle.heading4;

      // Determine if italic style (quote, subtitle)
      final isItalicStyle = paragraph.style == DocxParagraphStyle.quote ||
          paragraph.style == DocxParagraphStyle.subtitle;

      // Add prefix as regular text
      if (prefix.isNotEmpty) {
        segments.add(_TextSegment(prefix, '/F1'));
      }

      // Add runs with their formatting
      for (final run in paragraph.runs) {
        String fontRef;
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
        segments.add(_TextSegment(
          run.text,
          fontRef,
          color: run.color,
          underline: run.underline,
          strikethrough: run.strikethrough,
        ));
      }

      // Wrap text into lines
      final lines = _wrapTextSegments(segments, size);

      // Render each line
      for (int i = 0; i < lines.length; i++) {
        final line = lines[i];
        final lineText = line.map((s) => s.text).join();
        final lineWidth = _estimateTextWidth(lineText, size);

        // Calculate X position based on alignment and indent
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

        // Move to position for this line
        buffer.writeln('1 0 0 1 $xPos $currentY Tm');

        // Track current x position for underline/strikethrough
        var segmentX = xPos;

        // Render segments in this line
        for (final segment in line) {
          // Set color if specified
          if (segment.color != null) {
            final rgb = _hexToRgb(segment.color!);
            buffer.writeln('${rgb.$1} ${rgb.$2} ${rgb.$3} rg');
          } else {
            buffer.writeln('0 0 0 rg'); // Default black
          }

          buffer.writeln('${segment.fontRef} $size Tf');
          buffer.writeln('(${_escapePdfString(segment.text)}) Tj');

          // Calculate segment width for decorations
          final segmentWidth = _estimateTextWidth(segment.text, size);

          // Store decoration info for later (will be drawn outside BT/ET)
          if (segment.underline || segment.strikethrough) {
            // We need to end text block, draw line, then continue
            buffer.writeln('ET');
            buffer.writeln('q'); // Save graphics state

            // Set line color
            if (segment.color != null) {
              final rgb = _hexToRgb(segment.color!);
              buffer.writeln('${rgb.$1} ${rgb.$2} ${rgb.$3} RG');
            } else {
              buffer.writeln('0 0 0 RG'); // Default black stroke
            }

            buffer.writeln('0.5 w'); // Line width

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

            buffer.writeln('Q'); // Restore graphics state
            buffer.writeln('BT'); // Begin text again
            buffer.writeln('1 0 0 1 ${segmentX + segmentWidth} $currentY Tm');
          }

          segmentX += segmentWidth;
        }

        currentY -= lineHeight;
      }

      // Add paragraph spacing
      currentY -= size * 0.4;
    }

    buffer.writeln('ET'); // End text

    return buffer.toString();
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
    final spaceWidth = _estimateTextWidth(' ', size);

    for (final segment in segments) {
      final words = segment.text.split(' ');

      for (int i = 0; i < words.length; i++) {
        final word = words[i];
        if (word.isEmpty) continue;

        final wordWidth = _estimateTextWidth(word, size);
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

  double _estimateTextWidth(String text, int size) {
    // Approximate width: 0.5 * fontSize per character for proportional fonts
    return text.length * size * 0.5;
  }

  String _escapePdfString(String text) {
    return text
        .replaceAll('\\', '\\\\')
        .replaceAll('(', '\\(')
        .replaceAll(')', '\\)')
        .replaceAll('\n', '\\n')
        .replaceAll('\r', '\\r')
        .replaceAll('\t', '\\t');
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

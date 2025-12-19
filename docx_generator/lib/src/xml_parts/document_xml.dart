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

    for (final paragraph in document.paragraphs) {
      _writeParagraph(buffer, paragraph);
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
    final hasParagraphProps =
        hasStyle || hasAlignment || paragraph.pageBreakBefore;

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

  static void _writeRun(StringBuffer buffer, DocxRun run) {
    buffer.write('      <w:r>');

    // Run properties (formatting)
    if (run.hasFormatting) {
      buffer.write('<w:rPr>');
      if (run.bold) buffer.write('<w:b/>');
      if (run.italic) buffer.write('<w:i/>');
      if (run.underline) buffer.write('<w:u w:val="single"/>');
      if (run.strikethrough) buffer.write('<w:strike/>');
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
}

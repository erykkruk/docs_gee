import 'dart:convert';
import 'dart:typed_data';

import 'package:docs_gee/docs_gee.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

import 'support/docx_archive.dart';

/// Builds a PNG header with the given intrinsic size.
///
/// Only the signature and the IHDR chunk matter here: the library reads the
/// size straight out of the header and copies the bytes into the archive
/// without decoding pixels.
Uint8List pngBytes({required int width, required int height}) {
  final bytes = BytesBuilder();
  bytes.add([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
  bytes.add([0x00, 0x00, 0x00, 0x0D]); // IHDR length
  bytes.add(ascii.encode('IHDR'));
  bytes.add(_uint32(width));
  bytes.add(_uint32(height));
  bytes.add([0x08, 0x06, 0x00, 0x00, 0x00]); // bit depth, color type, etc.
  return bytes.toBytes();
}

/// Builds a JPEG with an APP0 segment followed by an SOF0 frame header.
Uint8List jpegBytes({required int width, required int height}) {
  final bytes = BytesBuilder();
  bytes.add([0xFF, 0xD8]); // SOI
  bytes.add([0xFF, 0xE0, 0x00, 0x04, 0x00, 0x00]); // APP0, length 4
  bytes.add([0xFF, 0xC0, 0x00, 0x11, 0x08]); // SOF0, length 17, precision 8
  bytes.add([(height >> 8) & 0xFF, height & 0xFF]);
  bytes.add([(width >> 8) & 0xFF, width & 0xFF]);
  bytes.add([0x03]); // component count
  return bytes.toBytes();
}

List<int> _uint32(int value) => [
      (value >> 24) & 0xFF,
      (value >> 16) & 0xFF,
      (value >> 8) & 0xFF,
      value & 0xFF,
    ];

void main() {
  group('DocxImage.detectFormat', () {
    test('detects a PNG signature', () {
      expect(
        DocxImage.detectFormat(pngBytes(width: 10, height: 10)),
        DocxImageFormat.png,
      );
    });

    test('detects a JPEG signature', () {
      expect(
        DocxImage.detectFormat(jpegBytes(width: 10, height: 10)),
        DocxImageFormat.jpeg,
      );
    });

    test('returns null for an unsupported format', () {
      final gif = Uint8List.fromList(ascii.encode('GIF89a-------'));
      expect(DocxImage.detectFormat(gif), isNull);
    });

    test('returns null for empty and truncated input', () {
      expect(DocxImage.detectFormat(Uint8List(0)), isNull);
      expect(DocxImage.detectFormat(Uint8List.fromList([0x89, 0x50])), isNull);
    });
  });

  group('DocxImage - intrinsic size', () {
    test('reads the size from a PNG IHDR chunk', () {
      final image = DocxImage(bytes: pngBytes(width: 640, height: 480));

      expect(image.intrinsicWidth, 640);
      expect(image.intrinsicHeight, 480);
      expect(image.width, 640);
      expect(image.height, 480);
      expect(image.format, DocxImageFormat.png);
    });

    test('reads the size from a JPEG frame header', () {
      final image = DocxImage(bytes: jpegBytes(width: 1024, height: 768));

      expect(image.intrinsicWidth, 1024);
      expect(image.intrinsicHeight, 768);
      expect(image.format, DocxImageFormat.jpeg);
    });

    test('skips over JPEG segments before the frame header', () {
      // The APP0 segment in the fixture must be walked over, not parsed as a
      // frame header, or the reported size is garbage.
      final image = DocxImage(bytes: jpegBytes(width: 300, height: 200));

      expect(image.intrinsicWidth, 300);
      expect(image.intrinsicHeight, 200);
    });
  });

  group('DocxImage - explicit sizing', () {
    test('honours both dimensions when given', () {
      final image = DocxImage(
        bytes: pngBytes(width: 800, height: 600),
        width: 200,
        height: 50,
      );

      expect(image.width, 200);
      expect(image.height, 50);
    });

    test('derives height from width, keeping the aspect ratio', () {
      final image = DocxImage(
        bytes: pngBytes(width: 800, height: 400),
        width: 200,
      );

      expect(image.width, 200);
      expect(image.height, 100);
    });

    test('derives width from height, keeping the aspect ratio', () {
      final image = DocxImage(
        bytes: pngBytes(width: 800, height: 400),
        height: 100,
      );

      expect(image.width, 200);
      expect(image.height, 100);
    });

    test('never scales an edge below one pixel', () {
      final image = DocxImage(
        bytes: pngBytes(width: 1000, height: 2),
        width: 1,
      );

      expect(image.height, greaterThanOrEqualTo(1));
    });

    test('converts pixels to EMU at 96 DPI', () {
      final image = DocxImage(bytes: pngBytes(width: 96, height: 48));

      expect(image.widthEmu, 96 * 9525);
      expect(image.heightEmu, 48 * 9525);
    });
  });

  group('DocxImage - rejected input', () {
    test('rejects an unsupported format', () {
      expect(
        () => DocxImage(bytes: Uint8List.fromList(ascii.encode('GIF89a--'))),
        throwsA(isA<DocxImageException>()),
      );
    });

    test('rejects a truncated PNG header', () {
      final truncated = Uint8List.fromList(
        pngBytes(width: 10, height: 10).sublist(0, 20),
      );
      expect(
        () => DocxImage(bytes: truncated),
        throwsA(isA<DocxImageException>()),
      );
    });

    test('rejects a JPEG with no frame header', () {
      final noFrame = Uint8List.fromList([
        0xFF, 0xD8, //
        0xFF, 0xE0, 0x00, 0x04, 0x00, 0x00,
      ]);
      expect(
        () => DocxImage(bytes: noFrame),
        throwsA(isA<DocxImageException>()),
      );
    });

    test('rejects a zero dimension in the header', () {
      expect(
        () => DocxImage(bytes: pngBytes(width: 0, height: 10)),
        throwsA(isA<DocxImageException>()),
      );
    });

    test('rejects non-positive size overrides', () {
      expect(
        () => DocxImage(bytes: pngBytes(width: 10, height: 10), width: 0),
        throwsA(isA<DocxImageException>()),
      );
      expect(
        () => DocxImage(bytes: pngBytes(width: 10, height: 10), height: -5),
        throwsA(isA<DocxImageException>()),
      );
    });
  });

  group('DOCX generation with images', () {
    test('writes the media part into the archive', () {
      final doc = DocxDocument()
        ..addImage(DocxImage(bytes: pngBytes(width: 100, height: 50)));

      final names = partNames(DocxGenerator().generate(doc));

      expect(names, contains('word/media/image1.png'));
    });

    test('declares the image extension in [Content_Types].xml', () {
      final doc = DocxDocument()
        ..addImage(DocxImage(bytes: pngBytes(width: 100, height: 50)));

      final types =
          partText(DocxGenerator().generate(doc), '[Content_Types].xml')!;

      expect(types, contains('Extension="png"'));
      expect(types, contains('image/png'));
    });

    test('declares one Default per format, not per image', () {
      final doc = DocxDocument()
        ..addImage(DocxImage(bytes: pngBytes(width: 10, height: 10)))
        ..addImage(DocxImage(bytes: pngBytes(width: 20, height: 20)))
        ..addImage(DocxImage(bytes: jpegBytes(width: 30, height: 30)));

      final types =
          partText(DocxGenerator().generate(doc), '[Content_Types].xml')!;

      expect('Extension="png"'.allMatches(types).length, 1);
      expect('Extension="jpeg"'.allMatches(types).length, 1);
    });

    test('adds an image relationship pointing at the media part', () {
      final doc = DocxDocument()
        ..addImage(DocxImage(bytes: pngBytes(width: 100, height: 50)));

      final rels = partText(
        DocxGenerator().generate(doc),
        'word/_rels/document.xml.rels',
      )!;

      expect(rels, contains('relationships/image'));
      expect(rels, contains('Target="media/image1.png"'));
    });

    test('references the relationship from the drawing markup', () {
      final doc = DocxDocument()
        ..addImage(DocxImage(bytes: pngBytes(width: 100, height: 50)));

      final bytes = DocxGenerator().generate(doc);
      final document = partText(bytes, 'word/document.xml')!;
      final rels = partText(bytes, 'word/_rels/document.xml.rels')!;

      expect(document, contains('<w:drawing>'));
      expect(document, contains('<wp:inline'));

      // The r:embed id must exist in the relationship file, or Word rejects
      // the document outright.
      final embed = RegExp(r'r:embed="(rId\d+)"').firstMatch(document);
      expect(embed, isNotNull);
      expect(rels, contains('Id="${embed!.group(1)}"'));
    });

    test('writes the rendered size in EMU', () {
      final doc = DocxDocument()
        ..addImage(DocxImage(
          bytes: pngBytes(width: 200, height: 100),
          width: 100,
        ));

      final document =
          partText(DocxGenerator().generate(doc), 'word/document.xml')!;

      expect(document, contains('cx="${100 * 9525}"'));
      expect(document, contains('cy="${50 * 9525}"'));
    });

    test('keeps media parts and relationship ids aligned for many images', () {
      final doc = DocxDocument()
        ..addImage(DocxImage(bytes: pngBytes(width: 10, height: 10)))
        ..addImage(DocxImage(bytes: jpegBytes(width: 20, height: 20)))
        ..addImage(DocxImage(bytes: pngBytes(width: 30, height: 30)));

      final bytes = DocxGenerator().generate(doc);
      final names = partNames(bytes);

      expect(names, contains('word/media/image1.png'));
      expect(names, contains('word/media/image2.jpeg'));
      expect(names, contains('word/media/image3.png'));
    });

    test('carries alt text into the drawing description', () {
      final doc = DocxDocument()
        ..addImage(DocxImage(
          bytes: pngBytes(width: 10, height: 10),
          altText: 'Quarterly revenue chart',
        ));

      final document =
          partText(DocxGenerator().generate(doc), 'word/document.xml')!;

      expect(document, contains('descr="Quarterly revenue chart"'));
    });

    test('escapes alt text so quotes cannot break the attribute', () {
      final doc = DocxDocument()
        ..addImage(DocxImage(
          bytes: pngBytes(width: 10, height: 10),
          altText: 'A "quoted" & <angled> caption',
        ));

      final document =
          partText(DocxGenerator().generate(doc), 'word/document.xml')!;

      expect(document, isNot(contains('descr="A "quoted"')));
      expect(document, contains('&quot;quoted&quot;'));
      expect(document, contains('&amp;'));
    });

    test('centers the image paragraph when asked', () {
      final doc = DocxDocument()
        ..addImage(DocxImage(
          bytes: pngBytes(width: 10, height: 10),
          alignment: DocxAlignment.center,
        ));

      final document =
          partText(DocxGenerator().generate(doc), 'word/document.xml')!;

      expect(document, contains('<w:jc w:val="center"/>'));
    });

    test('a document without images declares no image extensions', () {
      final doc = DocxDocument()
        ..addParagraph(DocxParagraph.text('No pictures here'));

      final bytes = DocxGenerator().generate(doc);

      expect(partNames(bytes).any((n) => n.startsWith('word/media/')), isFalse);
      expect(
        partText(bytes, '[Content_Types].xml'),
        isNot(contains('Extension="png"')),
      );
    });

    test('images are exposed on the document model', () {
      final doc = DocxDocument()
        ..addParagraph(DocxParagraph.text('Intro'))
        ..addImage(DocxImage(bytes: pngBytes(width: 10, height: 10)));

      expect(doc.images, hasLength(1));
      expect(doc.paragraphs, hasLength(1));
      expect(doc.content, hasLength(2));
    });
  });

  group('Archive integrity', () {
    test('every XML part parses, with images and a header present', () {
      // A malformed part is the failure mode that silently produces a file
      // Word refuses to open, so parse them all rather than trusting the
      // string assertions above.
      final doc = DocxDocument(
        header: DocxHeaderFooter.text('Report'),
        footer: DocxHeaderFooter.pageNumber(showTotal: true),
      )
        ..addParagraph(DocxParagraph.heading('Title', level: 1))
        ..addImage(DocxImage(
          bytes: pngBytes(width: 120, height: 60),
          altText: 'Chart & "notes"',
        ))
        ..addTable(DocxTable.simple(
          [
            ['A', 'B'],
          ],
          cellPadding: const DocxCellPadding.points(top: 4, left: 6),
        ));

      final bytes = DocxGenerator().generate(doc);

      for (final name in partNames(bytes)) {
        if (!name.endsWith('.xml') && !name.endsWith('.rels')) continue;
        expect(
          () => XmlDocument.parse(partText(bytes, name)!),
          returnsNormally,
          reason: name,
        );
      }
    });

    test('every r:id in document.xml resolves to a relationship', () {
      final doc = DocxDocument(
        header: DocxHeaderFooter.text('Report'),
        footer: DocxHeaderFooter.text('Footer'),
      )
        ..addParagraph(
          const DocxParagraph(
            runs: [DocxRun('link', hyperlink: 'https://example.com')],
          ),
        )
        ..addImage(DocxImage(bytes: pngBytes(width: 10, height: 10)))
        ..addImage(DocxImage(bytes: jpegBytes(width: 10, height: 10)));

      final bytes = DocxGenerator().generate(doc);
      final document = partText(bytes, 'word/document.xml')!;
      final rels = partText(bytes, 'word/_rels/document.xml.rels')!;

      final declared = RegExp(r'Id="(rId\d+)"')
          .allMatches(rels)
          .map((m) => m.group(1))
          .toSet();
      final referenced = RegExp(r'r:(?:id|embed)="(rId\d+)"')
          .allMatches(document)
          .map((m) => m.group(1))
          .toSet();

      expect(referenced, isNotEmpty);
      expect(declared, containsAll(referenced));
    });

    test('relationship ids are unique across parts of the document', () {
      final doc = DocxDocument(
        header: DocxHeaderFooter.text('Report'),
        footer: DocxHeaderFooter.text('Footer'),
      )
        ..addParagraph(DocxParagraph.bulletItem('A list, so numbering.xml'))
        ..addParagraph(
          const DocxParagraph(
            runs: [DocxRun('link', hyperlink: 'https://example.com')],
          ),
        )
        ..addImage(DocxImage(bytes: pngBytes(width: 10, height: 10)));

      final rels = partText(
        DocxGenerator().generate(doc),
        'word/_rels/document.xml.rels',
      )!;

      final ids =
          RegExp(r'Id="(rId\d+)"').allMatches(rels).map((m) => m.group(1));

      expect(ids.toSet().length, ids.length);
    });
  });
}

import 'package:docs_gee/docs_gee.dart';
import 'package:docs_gee/src/reader/style_resolver.dart';
import 'package:test/test.dart';

void main() {
  group('StyleResolver.resolveStyleId', () {
    test('resolves known style IDs', () {
      expect(
          StyleResolver.resolveStyleId('Normal'), DocxParagraphStyle.normal);
      expect(
          StyleResolver.resolveStyleId('Heading1'), DocxParagraphStyle.heading1);
      expect(
          StyleResolver.resolveStyleId('Heading2'), DocxParagraphStyle.heading2);
      expect(
          StyleResolver.resolveStyleId('Heading3'), DocxParagraphStyle.heading3);
      expect(
          StyleResolver.resolveStyleId('Heading4'), DocxParagraphStyle.heading4);
      expect(StyleResolver.resolveStyleId('Subtitle'),
          DocxParagraphStyle.subtitle);
      expect(
          StyleResolver.resolveStyleId('Caption'), DocxParagraphStyle.caption);
      expect(StyleResolver.resolveStyleId('Quote'), DocxParagraphStyle.quote);
      expect(StyleResolver.resolveStyleId('CodeBlock'),
          DocxParagraphStyle.codeBlock);
      expect(StyleResolver.resolveStyleId('Footnote'),
          DocxParagraphStyle.footnote);
      expect(StyleResolver.resolveStyleId('ListBullet'),
          DocxParagraphStyle.listBullet);
      expect(StyleResolver.resolveStyleId('ListDash'),
          DocxParagraphStyle.listDash);
      expect(StyleResolver.resolveStyleId('ListNumber'),
          DocxParagraphStyle.listNumber);
      expect(StyleResolver.resolveStyleId('ListNumberAlpha'),
          DocxParagraphStyle.listNumberAlpha);
      expect(StyleResolver.resolveStyleId('ListNumberRoman'),
          DocxParagraphStyle.listNumberRoman);
    });

    test('returns normal for unknown style IDs', () {
      expect(StyleResolver.resolveStyleId('UnknownStyle'),
          DocxParagraphStyle.normal);
      expect(StyleResolver.resolveStyleId('MyCustom'),
          DocxParagraphStyle.normal);
    });

    test('returns normal for null', () {
      expect(StyleResolver.resolveStyleId(null), DocxParagraphStyle.normal);
    });
  });

  group('StyleResolver.resolveNumId', () {
    test('resolves known numIds', () {
      expect(StyleResolver.resolveNumId(1), DocxParagraphStyle.listBullet);
      expect(StyleResolver.resolveNumId(2), DocxParagraphStyle.listNumber);
      expect(StyleResolver.resolveNumId(3), DocxParagraphStyle.listDash);
      expect(StyleResolver.resolveNumId(4), DocxParagraphStyle.listNumberAlpha);
      expect(
          StyleResolver.resolveNumId(5), DocxParagraphStyle.listNumberRoman);
    });

    test('falls back to bullet for unknown numId without numbering XML', () {
      expect(StyleResolver.resolveNumId(99), DocxParagraphStyle.listBullet);
    });
  });

  group('StyleResolver.resolveAlignment', () {
    test('resolves known alignments', () {
      expect(StyleResolver.resolveAlignment('left'), DocxAlignment.left);
      expect(StyleResolver.resolveAlignment('center'), DocxAlignment.center);
      expect(StyleResolver.resolveAlignment('right'), DocxAlignment.right);
      expect(StyleResolver.resolveAlignment('both'), DocxAlignment.justify);
      expect(StyleResolver.resolveAlignment('justify'), DocxAlignment.justify);
    });

    test('resolves start/end aliases', () {
      expect(StyleResolver.resolveAlignment('start'), DocxAlignment.left);
      expect(StyleResolver.resolveAlignment('end'), DocxAlignment.right);
    });

    test('defaults to left for unknown', () {
      expect(StyleResolver.resolveAlignment(null), DocxAlignment.left);
      expect(StyleResolver.resolveAlignment('unknown'), DocxAlignment.left);
    });
  });

  group('StyleResolver.resolveVerticalAlignment', () {
    test('resolves known values', () {
      expect(StyleResolver.resolveVerticalAlignment('top'),
          DocxVerticalAlignment.top);
      expect(StyleResolver.resolveVerticalAlignment('center'),
          DocxVerticalAlignment.center);
      expect(StyleResolver.resolveVerticalAlignment('bottom'),
          DocxVerticalAlignment.bottom);
    });

    test('defaults to top for unknown', () {
      expect(StyleResolver.resolveVerticalAlignment(null),
          DocxVerticalAlignment.top);
      expect(StyleResolver.resolveVerticalAlignment('middle'),
          DocxVerticalAlignment.top);
    });
  });
}

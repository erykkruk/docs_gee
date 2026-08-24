import 'dart:typed_data';

import 'docx_enums.dart';

/// Raster image formats that can be embedded in a DOCX document.
enum DocxImageFormat {
  /// PNG, embedded as `image/png`.
  png('png', 'image/png'),

  /// JPEG, embedded as `image/jpeg`.
  jpeg('jpeg', 'image/jpeg');

  const DocxImageFormat(this.extension, this.mimeType);

  /// Lowercase file extension used for the part name inside the archive.
  final String extension;

  /// MIME type declared in `[Content_Types].xml`.
  final String mimeType;
}

/// Thrown when image bytes cannot be embedded.
///
/// Either the format is not one of [DocxImageFormat], or the header is
/// truncated to the point where the intrinsic size cannot be read.
class DocxImageException implements Exception {
  /// Creates an exception describing why an image was rejected.
  const DocxImageException(this.message);

  /// Human-readable reason.
  final String message;

  @override
  String toString() => 'DocxImageException: $message';
}

/// An image embedded in a DOCX document.
///
/// The intrinsic pixel size is read straight from the PNG or JPEG header, so
/// a picture needs no explicit dimensions to come out at the right size:
///
/// ```dart
/// final doc = DocxDocument();
/// doc.addImage(DocxImage(bytes: pngBytes, altText: 'Quarterly chart'));
/// ```
///
/// Pass [width] and/or [height] (in pixels at 96 DPI) to scale it. Giving
/// only one keeps the aspect ratio:
///
/// ```dart
/// // 400px wide, height derived from the image's own proportions
/// DocxImage(bytes: pngBytes, width: 400);
/// ```
class DocxImage {
  /// Creates an image from raw PNG or JPEG bytes.
  ///
  /// Throws [DocxImageException] if [bytes] is not a supported format or the
  /// header is too short to carry the intrinsic size.
  factory DocxImage({
    required Uint8List bytes,
    int? width,
    int? height,
    String? altText,
    DocxAlignment alignment = DocxAlignment.left,
  }) {
    if (width != null && width <= 0) {
      throw const DocxImageException('width must be a positive pixel count');
    }
    if (height != null && height <= 0) {
      throw const DocxImageException('height must be a positive pixel count');
    }

    final format = detectFormat(bytes);
    if (format == null) {
      throw const DocxImageException(
        'unsupported image format: only PNG and JPEG can be embedded',
      );
    }
    final intrinsic = _readIntrinsicSize(bytes, format);

    final resolved = _resolveSize(
      intrinsicWidth: intrinsic.width,
      intrinsicHeight: intrinsic.height,
      width: width,
      height: height,
    );

    return DocxImage._(
      bytes: bytes,
      format: format,
      intrinsicWidth: intrinsic.width,
      intrinsicHeight: intrinsic.height,
      width: resolved.width,
      height: resolved.height,
      altText: altText,
      alignment: alignment,
    );
  }

  const DocxImage._({
    required this.bytes,
    required this.format,
    required this.intrinsicWidth,
    required this.intrinsicHeight,
    required this.width,
    required this.height,
    required this.altText,
    required this.alignment,
  });

  /// Raw encoded image bytes, written to the archive as-is.
  final Uint8List bytes;

  /// Detected image format.
  final DocxImageFormat format;

  /// Width in pixels as stored in the file header.
  final int intrinsicWidth;

  /// Height in pixels as stored in the file header.
  final int intrinsicHeight;

  /// Rendered width in pixels at 96 DPI.
  final int width;

  /// Rendered height in pixels at 96 DPI.
  final int height;

  /// Alternative text, exposed to screen readers as the picture description.
  final String? altText;

  /// Horizontal placement of the image within the page.
  final DocxAlignment alignment;

  /// English Metric Units per pixel at 96 DPI (914400 EMU per inch).
  static const int emuPerPixel = 9525;

  /// Rendered width in EMU, the unit the drawing markup uses.
  int get widthEmu => width * emuPerPixel;

  /// Rendered height in EMU, the unit the drawing markup uses.
  int get heightEmu => height * emuPerPixel;

  /// Returns the format of [bytes], or null when it is neither PNG nor JPEG.
  static DocxImageFormat? detectFormat(Uint8List bytes) {
    if (bytes.length >= 8 &&
        bytes[0] == 0x89 &&
        bytes[1] == 0x50 &&
        bytes[2] == 0x4E &&
        bytes[3] == 0x47 &&
        bytes[4] == 0x0D &&
        bytes[5] == 0x0A &&
        bytes[6] == 0x1A &&
        bytes[7] == 0x0A) {
      return DocxImageFormat.png;
    }
    if (bytes.length >= 3 &&
        bytes[0] == 0xFF &&
        bytes[1] == 0xD8 &&
        bytes[2] == 0xFF) {
      return DocxImageFormat.jpeg;
    }
    return null;
  }

  /// Reads the pixel size from a PNG IHDR chunk or a JPEG SOF marker.
  static ({int width, int height}) _readIntrinsicSize(
    Uint8List bytes,
    DocxImageFormat format,
  ) {
    return switch (format) {
      DocxImageFormat.png => _readPngSize(bytes),
      DocxImageFormat.jpeg => _readJpegSize(bytes),
    };
  }

  /// PNG stores width and height as big-endian uint32 at offsets 16 and 20,
  /// inside the IHDR chunk that the spec requires to come first.
  static ({int width, int height}) _readPngSize(Uint8List bytes) {
    if (bytes.length < 24) {
      throw const DocxImageException(
        'truncated PNG: header too short to contain an IHDR chunk',
      );
    }
    final width = _readUint32(bytes, 16);
    final height = _readUint32(bytes, 20);
    if (width == 0 || height == 0) {
      throw const DocxImageException('PNG reports a zero dimension');
    }
    return (width: width, height: height);
  }

  /// JPEG keeps the size in a Start Of Frame segment, so the marker chain has
  /// to be walked until one shows up.
  static ({int width, int height}) _readJpegSize(Uint8List bytes) {
    var offset = 2; // skip the SOI marker
    while (offset + 3 < bytes.length) {
      if (bytes[offset] != 0xFF) {
        offset++;
        continue;
      }
      final marker = bytes[offset + 1];
      // Padding and standalone markers carry no length field.
      if (marker == 0xFF ||
          marker == 0x01 ||
          (marker >= 0xD0 && marker <= 0xD9)) {
        offset += 2;
        continue;
      }
      final segmentLength = (bytes[offset + 2] << 8) | bytes[offset + 3];
      if (segmentLength < 2) {
        throw const DocxImageException(
            'malformed JPEG: invalid segment length');
      }
      // SOF0-SOF15 hold the frame header; DHT (C4), JPG (C8) and DAC (CC)
      // share the range but are not frame headers.
      final isFrameHeader = marker >= 0xC0 &&
          marker <= 0xCF &&
          marker != 0xC4 &&
          marker != 0xC8 &&
          marker != 0xCC;
      if (isFrameHeader) {
        if (offset + 9 >= bytes.length) {
          throw const DocxImageException(
            'truncated JPEG: frame header is incomplete',
          );
        }
        final height = (bytes[offset + 5] << 8) | bytes[offset + 6];
        final width = (bytes[offset + 7] << 8) | bytes[offset + 8];
        if (width == 0 || height == 0) {
          throw const DocxImageException('JPEG reports a zero dimension');
        }
        return (width: width, height: height);
      }
      offset += 2 + segmentLength;
    }
    throw const DocxImageException(
      'malformed JPEG: no frame header found before the end of the data',
    );
  }

  /// Applies the caller's size overrides, preserving the aspect ratio when
  /// only one edge is given.
  static ({int width, int height}) _resolveSize({
    required int intrinsicWidth,
    required int intrinsicHeight,
    required int? width,
    required int? height,
  }) {
    if (width != null && height != null) {
      return (width: width, height: height);
    }
    if (width != null) {
      final scaled = (intrinsicHeight * width / intrinsicWidth).round();
      return (width: width, height: scaled < 1 ? 1 : scaled);
    }
    if (height != null) {
      final scaled = (intrinsicWidth * height / intrinsicHeight).round();
      return (width: scaled < 1 ? 1 : scaled, height: height);
    }
    return (width: intrinsicWidth, height: intrinsicHeight);
  }

  static int _readUint32(Uint8List bytes, int offset) {
    return (bytes[offset] << 24) |
        (bytes[offset + 1] << 16) |
        (bytes[offset + 2] << 8) |
        bytes[offset + 3];
  }
}

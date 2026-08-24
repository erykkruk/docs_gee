import 'dart:convert';
import 'dart:typed_data';

import 'package:archive/archive.dart';

/// Reads one part out of a generated DOCX archive as text.
///
/// Returns null when the archive has no such part, which is what most
/// "should not be written" assertions check for.
String? partText(Uint8List docx, String path) {
  final archive = ZipDecoder().decodeBytes(docx);
  for (final file in archive.files) {
    if (file.name == path) {
      return utf8.decode(file.content as List<int>);
    }
  }
  return null;
}

/// Lists the part names inside a generated DOCX archive.
List<String> partNames(Uint8List docx) {
  return ZipDecoder().decodeBytes(docx).files.map((f) => f.name).toList();
}

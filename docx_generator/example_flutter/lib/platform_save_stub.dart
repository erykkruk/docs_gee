/// Stub implementation for unsupported platforms.
library;

import 'dart:typed_data';

Future<void> saveDocxFile(Uint8List bytes, String filename) async {
  throw UnsupportedError('Platform not supported');
}

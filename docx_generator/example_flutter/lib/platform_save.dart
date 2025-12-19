/// Platform-specific file saving implementation.
///
/// This file provides a unified interface for saving DOCX files
/// across all platforms (web, mobile, desktop).
library;

import 'dart:typed_data';

import 'package:flutter/foundation.dart';

import 'platform_save_stub.dart'
    if (dart.library.html) 'platform_save_web.dart'
    if (dart.library.io) 'platform_save_io.dart' as platform;

/// Saves a DOCX file using the appropriate platform mechanism.
///
/// - Web: Downloads the file through the browser
/// - Mobile: Saves to app documents directory
/// - Desktop: Saves to user-specified location
Future<void> saveDocxFile(Uint8List bytes, String filename) async {
  await platform.saveDocxFile(bytes, filename);
}

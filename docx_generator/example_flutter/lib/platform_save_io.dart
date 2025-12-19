/// Native (mobile/desktop) implementation for file saving.
///
/// Saves files to the appropriate location based on platform.
library;

import 'dart:io';
import 'dart:typed_data';

import 'package:flutter/foundation.dart';

/// Saves a DOCX file to the file system.
///
/// - Desktop: Saves to current directory
/// - Mobile: Saves to app documents directory (requires path_provider for production)
Future<void> saveDocxFile(Uint8List bytes, String filename) async {
  final String path;

  if (Platform.isAndroid || Platform.isIOS) {
    // For a production app, use path_provider package:
    // final dir = await getApplicationDocumentsDirectory();
    // path = '${dir.path}/$filename';

    // For this example, we'll use a temporary solution
    path = '/tmp/$filename';
  } else {
    // Desktop platforms
    path = filename;
  }

  final file = File(path);
  await file.writeAsBytes(bytes);

  if (kDebugMode) {
    print('Saved to: ${file.absolute.path}');
  }
}

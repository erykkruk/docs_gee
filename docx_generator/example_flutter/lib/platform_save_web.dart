/// Web implementation for file saving.
///
/// Uses browser download API to save files.
library;

// ignore: avoid_web_libraries_in_flutter
import 'dart:html' as html;
import 'dart:typed_data';

/// Downloads a DOCX file through the browser.
Future<void> saveDocxFile(Uint8List bytes, String filename) async {
  final blob = html.Blob(
    [bytes],
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  );
  final url = html.Url.createObjectUrlFromBlob(blob);

  html.AnchorElement(href: url)
    ..setAttribute('download', filename)
    ..click();

  html.Url.revokeObjectUrl(url);
}

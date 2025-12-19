import 'dart:io';
import 'dart:typed_data';

/// File saving utilities for non-Web platforms.
///
/// **Note:** This file uses `dart:io` and should NOT be imported on Web.
/// For Web platform, use the bytes from [generate] and handle saving
/// with browser APIs (e.g., `html.AnchorElement` or a file saver package).
///
/// Example usage:
/// ```dart
/// // Only import on non-Web platforms:
/// import 'package:docs_gee/src/file_saver_io.dart';
///
/// final doc = DocxDocument();
/// doc.addParagraph(DocxParagraph.text('Hello'));
///
/// final generator = DocxGenerator();
/// await saveToFile(generator.generate(doc), 'output.docx');
/// ```
class FileSaverIO {
  FileSaverIO._();

  /// Saves bytes to a file at the specified path.
  ///
  /// Returns the path where the file was saved.
  static Future<String> saveToFile(Uint8List bytes, String filePath) async {
    await File(filePath).writeAsBytes(bytes);
    return filePath;
  }

  /// Generates a document and saves it to a file.
  ///
  /// [bytes] - the document bytes from generator.generate()
  /// [filePath] - optional path. If null, uses [defaultPath].
  /// [defaultPath] - fallback path when [filePath] is null.
  ///
  /// Returns the actual file path where the document was saved.
  static Future<String> saveDocument(
    Uint8List bytes, {
    String? filePath,
    String defaultPath = 'document',
  }) async {
    final path = filePath ?? defaultPath;
    await File(path).writeAsBytes(bytes);
    return path;
  }
}

/// Extension methods for saving document bytes to files.
///
/// Only available on non-Web platforms.
extension Uint8ListFileSaver on Uint8List {
  /// Saves these bytes to a file at the specified path.
  Future<String> saveToFile(String filePath) async {
    await File(filePath).writeAsBytes(this);
    return filePath;
  }
}

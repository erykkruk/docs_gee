/// Base exception for DOCX reading errors.
sealed class DocxReaderException implements Exception {
  const DocxReaderException(this.message);

  /// Human-readable error description.
  final String message;

  @override
  String toString() => '$runtimeType: $message';
}

/// Thrown when the input bytes are not a valid ZIP archive.
class InvalidDocxArchiveException extends DocxReaderException {
  const InvalidDocxArchiveException([super.message = 'Invalid DOCX archive']);
}

/// Thrown when a required DOCX part (e.g. word/document.xml) is missing.
class MissingDocxPartException extends DocxReaderException {
  const MissingDocxPartException(String partName)
      : super('Missing required DOCX part: $partName');
}

/// Thrown when XML content inside the DOCX is malformed.
class InvalidDocxXmlException extends DocxReaderException {
  const InvalidDocxXmlException([super.message = 'Malformed XML in DOCX']);
}

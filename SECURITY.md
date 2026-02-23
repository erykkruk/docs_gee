# Security Policy

## Supported Versions

| Version | Supported |
|---------|-----------|
| Latest  | Yes       |
| < Latest | No       |

We recommend always using the latest version of docx_generator.

## Reporting a Vulnerability

If you discover a security vulnerability in docx_generator, please report it responsibly.

**Do NOT open a public GitHub issue for security vulnerabilities.**

Instead, please report security issues by emailing the maintainers directly. You can find contact information in the repository's profile or `pubspec.yaml`.

### What to Include

- Description of the vulnerability
- Steps to reproduce
- Potential impact
- Suggested fix (if any)

### Response Timeline

- **Acknowledgment**: Within 48 hours
- **Assessment**: Within 1 week
- **Fix/Disclosure**: Coordinated with reporter

## Scope

This policy applies to the `docx_generator` Dart package. Since this is a document generation library that produces DOCX files, the main security considerations include:

- XML injection in generated documents
- Path traversal in archive generation
- Denial of service through resource exhaustion

Thank you for helping keep docx_generator safe!

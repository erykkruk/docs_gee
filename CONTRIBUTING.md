# Contributing to docx_generator

Thank you for your interest in contributing to docx_generator! This guide will help you get started.

## How to Contribute

### Reporting Bugs

1. Check [existing issues](https://github.com/ArturoTaworworski/docs_gee/issues) to avoid duplicates
2. Use the **Bug Report** issue template
3. Include: Dart version, OS, minimal reproduction code, expected vs actual behavior

### Suggesting Features

1. Open an issue using the **Feature Request** template
2. Describe the problem you're trying to solve
3. Propose a solution and any alternatives you've considered

### Submitting Pull Requests

1. Fork the repository
2. Create a feature branch from `main`: `git checkout -b feature/your-feature`
3. Make your changes
4. Ensure all checks pass (see below)
5. Submit a PR using the pull request template

## Development Setup

```bash
cd docx_generator
dart pub get
```

## Before Submitting

Run these checks locally:

```bash
# Analyze code
dart analyze

# Format code
dart format .

# Run example to verify output
dart run example/example.dart
```

## Code Style

- Follow [Effective Dart](https://dart.dev/effective-dart) guidelines
- Use `dart format` for consistent formatting
- No `dynamic` types unless absolutely necessary
- Add doc comments for public APIs
- Keep functions focused on a single responsibility

## Commit Messages

Use [Conventional Commits](https://www.conventionalcommits.org/):

```
feat: add table cell merging support
fix: correct heading font size in styles.xml
docs: update API documentation
test: add tests for paragraph alignment
```

## Project Structure

```
docx_generator/
├── lib/
│   ├── docx_generator.dart          # Public API exports
│   └── src/
│       ├── docx_generator.dart      # Main generator class
│       ├── models/                  # Data models
│       └── xml_parts/              # XML generation
├── example/
│   └── example.dart                # Usage examples
└── test/                           # Tests
```

## Questions?

Open a [discussion](https://github.com/ArturoTaworworski/docs_gee/issues) or reach out via an issue. We're happy to help!

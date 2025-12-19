/// Flutter example demonstrating cross-platform DOCX generation.
///
/// This example works on:
/// - iOS
/// - Android
/// - Web
/// - macOS
/// - Windows
/// - Linux
library;

import 'dart:typed_data';

import 'package:docs_gee/docs_gee.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';

import 'platform_save.dart';

void main() {
  runApp(const DocxGeneratorExample());
}

class DocxGeneratorExample extends StatelessWidget {
  const DocxGeneratorExample({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'DOCX Generator Demo',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.blue),
        useMaterial3: true,
      ),
      home: const DocxGeneratorPage(),
    );
  }
}

class DocxGeneratorPage extends StatefulWidget {
  const DocxGeneratorPage({super.key});

  @override
  State<DocxGeneratorPage> createState() => _DocxGeneratorPageState();
}

class _DocxGeneratorPageState extends State<DocxGeneratorPage> {
  final _titleController = TextEditingController(text: 'My Document');
  final _authorController = TextEditingController(text: 'Flutter User');
  final _contentController = TextEditingController(
    text:
        'This is a sample document generated using the docx_generator library.\n\n'
        'You can edit this text and generate a new document.',
  );

  String _fontName = 'Times New Roman';
  int _fontSize = 24;
  bool _isGenerating = false;
  String? _statusMessage;

  final _fonts = [
    'Times New Roman',
    'Arial',
    'Calibri',
    'Georgia',
    'Verdana',
  ];

  @override
  void dispose() {
    _titleController.dispose();
    _authorController.dispose();
    _contentController.dispose();
    super.dispose();
  }

  Future<void> _generateAndSave() async {
    setState(() {
      _isGenerating = true;
      _statusMessage = null;
    });

    try {
      final stopwatch = Stopwatch()..start();
      final bytes = _generateDocument();
      stopwatch.stop();

      await saveDocxFile(bytes, 'document.docx');

      setState(() {
        _statusMessage = 'Generated ${bytes.length} bytes '
            'in ${stopwatch.elapsedMilliseconds}ms';
      });
    } catch (e) {
      setState(() {
        _statusMessage = 'Error: $e';
      });
    } finally {
      setState(() {
        _isGenerating = false;
      });
    }
  }

  Uint8List _generateDocument() {
    final doc = DocxDocument(
      title: _titleController.text,
      author: _authorController.text,
    );

    // Add title
    doc.addParagraph(DocxParagraph.heading(
      _titleController.text,
      level: 1,
      alignment: DocxAlignment.center,
    ));

    // Add author
    doc.addParagraph(DocxParagraph.text(
      'By: ${_authorController.text}',
      alignment: DocxAlignment.center,
    ));

    doc.addParagraph(DocxParagraph.text(''));

    // Add content paragraphs
    final paragraphs = _contentController.text.split('\n');
    for (final text in paragraphs) {
      if (text.trim().isEmpty) {
        doc.addParagraph(DocxParagraph.text(''));
      } else {
        doc.addParagraph(DocxParagraph.text(
          text,
          alignment: DocxAlignment.justify,
        ));
      }
    }

    // Add demo section with formatting
    doc.addParagraph(DocxParagraph.heading(
      'Formatting Examples',
      level: 2,
      pageBreakBefore: true,
    ));

    doc.addParagraph(DocxParagraph(
      runs: [
        const DocxRun('This shows '),
        const DocxRun('bold', bold: true),
        const DocxRun(', '),
        const DocxRun('italic', italic: true),
        const DocxRun(', '),
        const DocxRun('underline', underline: true),
        const DocxRun(', and '),
        const DocxRun('strikethrough', strikethrough: true),
        const DocxRun(' formatting.'),
      ],
    ));

    doc.addParagraph(DocxParagraph.heading('Features List', level: 2));
    doc.addParagraph(DocxParagraph.bulletItem('Cross-platform support'));
    doc.addParagraph(DocxParagraph.bulletItem('Rich text formatting'));
    doc.addParagraph(DocxParagraph.bulletItem('Multiple heading levels'));
    doc.addParagraph(DocxParagraph.bulletItem('Bullet and numbered lists'));

    // Generate
    final generator = DocxGenerator(
      fontName: _fontName,
      fontSize: _fontSize,
    );

    return generator.generate(doc);
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('DOCX Generator'),
        backgroundColor: Theme.of(context).colorScheme.inversePrimary,
      ),
      body: SingleChildScrollView(
        padding: const EdgeInsets.all(16),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            // Document metadata
            Card(
              child: Padding(
                padding: const EdgeInsets.all(16),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text(
                      'Document Settings',
                      style: Theme.of(context).textTheme.titleMedium,
                    ),
                    const SizedBox(height: 16),
                    TextField(
                      controller: _titleController,
                      decoration: const InputDecoration(
                        labelText: 'Document Title',
                        border: OutlineInputBorder(),
                      ),
                    ),
                    const SizedBox(height: 12),
                    TextField(
                      controller: _authorController,
                      decoration: const InputDecoration(
                        labelText: 'Author',
                        border: OutlineInputBorder(),
                      ),
                    ),
                    const SizedBox(height: 12),
                    Row(
                      children: [
                        Expanded(
                          child: DropdownButtonFormField<String>(
                            value: _fontName,
                            decoration: const InputDecoration(
                              labelText: 'Font',
                              border: OutlineInputBorder(),
                            ),
                            items: _fonts.map((font) {
                              return DropdownMenuItem(
                                value: font,
                                child: Text(font),
                              );
                            }).toList(),
                            onChanged: (value) {
                              if (value != null) {
                                setState(() => _fontName = value);
                              }
                            },
                          ),
                        ),
                        const SizedBox(width: 12),
                        Expanded(
                          child: DropdownButtonFormField<int>(
                            value: _fontSize,
                            decoration: const InputDecoration(
                              labelText: 'Font Size',
                              border: OutlineInputBorder(),
                            ),
                            items: [20, 22, 24, 28, 32].map((size) {
                              return DropdownMenuItem(
                                value: size,
                                child: Text('${size ~/ 2}pt'),
                              );
                            }).toList(),
                            onChanged: (value) {
                              if (value != null) {
                                setState(() => _fontSize = value);
                              }
                            },
                          ),
                        ),
                      ],
                    ),
                  ],
                ),
              ),
            ),
            const SizedBox(height: 16),

            // Content input
            Card(
              child: Padding(
                padding: const EdgeInsets.all(16),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text(
                      'Document Content',
                      style: Theme.of(context).textTheme.titleMedium,
                    ),
                    const SizedBox(height: 16),
                    TextField(
                      controller: _contentController,
                      maxLines: 8,
                      decoration: const InputDecoration(
                        hintText: 'Enter your document content...',
                        border: OutlineInputBorder(),
                      ),
                    ),
                  ],
                ),
              ),
            ),
            const SizedBox(height: 16),

            // Generate button
            FilledButton.icon(
              onPressed: _isGenerating ? null : _generateAndSave,
              icon: _isGenerating
                  ? const SizedBox(
                      width: 20,
                      height: 20,
                      child: CircularProgressIndicator(
                        strokeWidth: 2,
                        color: Colors.white,
                      ),
                    )
                  : const Icon(Icons.description),
              label: Text(_isGenerating ? 'Generating...' : 'Generate DOCX'),
            ),

            // Status message
            if (_statusMessage != null) ...[
              const SizedBox(height: 16),
              Card(
                color: _statusMessage!.startsWith('Error')
                    ? Theme.of(context).colorScheme.errorContainer
                    : Theme.of(context).colorScheme.primaryContainer,
                child: Padding(
                  padding: const EdgeInsets.all(16),
                  child: Row(
                    children: [
                      Icon(
                        _statusMessage!.startsWith('Error')
                            ? Icons.error_outline
                            : Icons.check_circle_outline,
                      ),
                      const SizedBox(width: 12),
                      Expanded(child: Text(_statusMessage!)),
                    ],
                  ),
                ),
              ),
            ],

            const SizedBox(height: 16),

            // Platform info
            Card(
              child: Padding(
                padding: const EdgeInsets.all(16),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text(
                      'Platform Info',
                      style: Theme.of(context).textTheme.titleMedium,
                    ),
                    const SizedBox(height: 8),
                    Text('Running on: ${_getPlatformName()}'),
                    Text('Web: ${kIsWeb ? 'Yes' : 'No'}'),
                  ],
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }

  String _getPlatformName() {
    if (kIsWeb) return 'Web';
    switch (defaultTargetPlatform) {
      case TargetPlatform.android:
        return 'Android';
      case TargetPlatform.iOS:
        return 'iOS';
      case TargetPlatform.macOS:
        return 'macOS';
      case TargetPlatform.windows:
        return 'Windows';
      case TargetPlatform.linux:
        return 'Linux';
      case TargetPlatform.fuchsia:
        return 'Fuchsia';
    }
  }
}

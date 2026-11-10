import 'dart:io';
import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:http/http.dart' as http;

void main() {
  runApp(const EditorPdfApp());
}

class EditorPdfApp extends StatelessWidget {
  const EditorPdfApp({super.key});

  @override
  Widget build(BuildContext context) {
    const indigo = Color(0xFF3F51B5);
    return MaterialApp(
      title: 'PDF a Word',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        useMaterial3: true,
        fontFamily: 'Segoe UI',
        colorScheme: ColorScheme.fromSeed(seedColor: indigo),
        scaffoldBackgroundColor: const Color(0xFFF4F5F9),
      ),
      home: const HomePage(),
    );
  }
}

class HomePage extends StatefulWidget {
  const HomePage({super.key});

  @override
  State<HomePage> createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> {
  final _urlController = TextEditingController();

  String? _pdfPath; // ruta local del PDF listo para convertir
  String _status = 'Elige un PDF o pega una URL para empezar.';
  bool _busy = false;
  String? _resultPath; // .docx generado

  @override
  void dispose() {
    _urlController.dispose();
    super.dispose();
  }

  void _set(String msg, {bool busy = false}) {
    setState(() {
      _status = msg;
      _busy = busy;
    });
  }

  Future<void> _pickPdf() async {
    final result = await FilePicker.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['pdf'],
      dialogTitle: 'Selecciona un PDF',
    );
    if (result == null || result.files.single.path == null) return;
    setState(() {
      _pdfPath = result.files.single.path;
      _resultPath = null;
      _status = 'PDF listo: ${_fileName(_pdfPath!)}';
    });
  }

  Future<void> _downloadFromUrl() async {
    final url = _urlController.text.trim();
    if (url.isEmpty) return;
    if (!url.toLowerCase().startsWith('http')) {
      _set('La URL debe empezar con http:// o https://');
      return;
    }
    _set('Descargando PDF...', busy: true);
    try {
      final resp = await http.get(Uri.parse(url));
      if (resp.statusCode != 200) {
        _set('No se pudo descargar (código ${resp.statusCode}).');
        return;
      }
      final tempDir = Directory.systemTemp;
      var name = Uri.parse(url).pathSegments.isNotEmpty
          ? Uri.parse(url).pathSegments.last
          : 'descarga.pdf';
      if (!name.toLowerCase().endsWith('.pdf')) name = 'descarga.pdf';
      final file = File('${tempDir.path}\\$name');
      await file.writeAsBytes(resp.bodyBytes);
      setState(() {
        _pdfPath = file.path;
        _resultPath = null;
        _busy = false;
        _status = 'PDF descargado: $name';
      });
    } catch (e) {
      _set('Error al descargar: $e');
    }
  }

  Future<void> _convertToWord() async {
    final input = _pdfPath;
    if (input == null) return;

    // Elegir dónde guardar el .docx
    final suggested = '${_fileName(input).replaceAll(
      RegExp(r'\.pdf$', caseSensitive: false),
      '',
    )}.docx';
    final outPath = await FilePicker.saveFile(
      dialogTitle: 'Guardar Word como...',
      fileName: suggested,
      type: FileType.custom,
      allowedExtensions: ['docx'],
    );
    if (outPath == null) return;

    _set('Convirtiendo con Word... (puede tardar unos segundos)', busy: true);
    try {
      final script = await _writeConversionScript();
      final result = await Process.run(
        'powershell',
        [
          '-NoProfile',
          '-ExecutionPolicy',
          'Bypass',
          '-File',
          script.path,
          input,
          outPath,
        ],
      );
      final out = (result.stdout.toString() + result.stderr.toString()).trim();
      if (result.exitCode == 0 && File(outPath).existsSync()) {
        setState(() {
          _busy = false;
          _resultPath = outPath;
          _status = '¡Listo! Word creado: ${_fileName(outPath)}';
        });
      } else {
        _set('No se pudo convertir.\n$out');
      }
    } catch (e) {
      _set('Error en la conversión: $e');
    }
  }

  /// Escribe el script PowerShell que automatiza Word para convertir PDF -> DOCX.
  Future<File> _writeConversionScript() async {
    const ps = r'''
param([Parameter(Mandatory=$true)][string]$InPath,
      [Parameter(Mandatory=$true)][string]$OutPath)
$ErrorActionPreference = 'Stop'
$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    # ConfirmConversions=$false evita el aviso "Word convertira tu PDF"; ReadOnly=$true
    $doc = $word.Documents.Open($InPath, $false, $true)
    # 16 = wdFormatDocumentDefault (.docx)
    $doc.SaveAs2($OutPath, 16)
    Write-Output 'OK'
} catch {
    Write-Output ('ERROR: ' + $_.Exception.Message)
    exit 1
} finally {
    if ($doc -ne $null) { try { $doc.Close(0) } catch {} }
    if ($word -ne $null) { try { $word.Quit() } catch {} }
}
''';
    final file = File('${Directory.systemTemp.path}\\pdf2word_convert.ps1');
    await file.writeAsString(ps);
    return file;
  }

  void _openResult() {
    if (_resultPath != null) {
      Process.run('cmd', ['/c', 'start', '', _resultPath!]);
    }
  }

  void _openFolder() {
    if (_resultPath != null) {
      Process.run('explorer', ['/select,', _resultPath!]);
    }
  }

  String _fileName(String path) => path.split(RegExp(r'[\\/]')).last;

  @override
  Widget build(BuildContext context) {
    final indigo = Theme.of(context).colorScheme.primary;
    return Scaffold(
      appBar: AppBar(
        backgroundColor: indigo,
        foregroundColor: Colors.white,
        title: const Text('Convertir PDF a Word'),
        centerTitle: false,
      ),
      body: Center(
        child: ConstrainedBox(
          constraints: const BoxConstraints(maxWidth: 620),
          child: SingleChildScrollView(
            padding: const EdgeInsets.all(28),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.stretch,
              children: [
                _card(
                  title: '1. Elige tu PDF',
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.stretch,
                    children: [
                      FilledButton.icon(
                        onPressed: _busy ? null : _pickPdf,
                        icon: const Icon(Icons.folder_open),
                        label: const Text('Elegir archivo PDF'),
                      ),
                      const SizedBox(height: 16),
                      const Row(children: [
                        Expanded(child: Divider()),
                        Padding(
                          padding: EdgeInsets.symmetric(horizontal: 8),
                          child: Text('o desde una URL'),
                        ),
                        Expanded(child: Divider()),
                      ]),
                      const SizedBox(height: 12),
                      Row(children: [
                        Expanded(
                          child: TextField(
                            controller: _urlController,
                            decoration: const InputDecoration(
                              hintText: 'https://...archivo.pdf',
                              border: OutlineInputBorder(),
                              isDense: true,
                            ),
                          ),
                        ),
                        const SizedBox(width: 8),
                        OutlinedButton(
                          onPressed: _busy ? null : _downloadFromUrl,
                          child: const Text('Descargar'),
                        ),
                      ]),
                    ],
                  ),
                ),
                const SizedBox(height: 16),
                _card(
                  title: '2. Convertir',
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.stretch,
                    children: [
                      FilledButton.icon(
                        onPressed:
                            (_busy || _pdfPath == null) ? null : _convertToWord,
                        style: FilledButton.styleFrom(
                          padding: const EdgeInsets.symmetric(vertical: 16),
                        ),
                        icon: const Icon(Icons.description),
                        label: const Text('Convertir a Word (.docx)'),
                      ),
                      if (_resultPath != null) ...[
                        const SizedBox(height: 12),
                        Row(children: [
                          Expanded(
                            child: OutlinedButton.icon(
                              onPressed: _openResult,
                              icon: const Icon(Icons.open_in_new),
                              label: const Text('Abrir Word'),
                            ),
                          ),
                          const SizedBox(width: 8),
                          Expanded(
                            child: OutlinedButton.icon(
                              onPressed: _openFolder,
                              icon: const Icon(Icons.folder),
                              label: const Text('Ver carpeta'),
                            ),
                          ),
                        ]),
                      ],
                    ],
                  ),
                ),
                const SizedBox(height: 20),
                Container(
                  padding: const EdgeInsets.all(14),
                  decoration: BoxDecoration(
                    color: Colors.white,
                    borderRadius: BorderRadius.circular(10),
                    border: Border.all(color: const Color(0xFFE0E2EC)),
                  ),
                  child: Row(children: [
                    if (_busy)
                      const SizedBox(
                        width: 18,
                        height: 18,
                        child: CircularProgressIndicator(strokeWidth: 2),
                      )
                    else
                      Icon(Icons.info_outline, size: 18, color: indigo),
                    const SizedBox(width: 10),
                    Expanded(child: Text(_status)),
                  ]),
                ),
                const SizedBox(height: 12),
                const Text(
                  'La conversión usa tu Microsoft Word instalado. '
                  'Todo ocurre en tu equipo: nada se sube a internet.',
                  style: TextStyle(fontSize: 12, color: Colors.black54),
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }

  Widget _card({required String title, required Widget child}) {
    return Container(
      padding: const EdgeInsets.all(20),
      decoration: BoxDecoration(
        color: Colors.white,
        borderRadius: BorderRadius.circular(14),
        boxShadow: [
          BoxShadow(
            color: Colors.black.withValues(alpha: 0.05),
            blurRadius: 10,
            offset: const Offset(0, 3),
          ),
        ],
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Text(title,
              style: const TextStyle(
                  fontSize: 16, fontWeight: FontWeight.w600)),
          const SizedBox(height: 14),
          child,
        ],
      ),
    );
  }
}

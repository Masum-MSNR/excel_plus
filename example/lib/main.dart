import 'dart:io';

import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';

import 'tests/all_tests.dart';
import 'tests/test_case.dart';

void main() {
  runApp(const ExcelPlusTestApp());
}

class ExcelPlusTestApp extends StatelessWidget {
  const ExcelPlusTestApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'excel_plus Test',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        colorSchemeSeed: Colors.blue,
        useMaterial3: true,
      ),
      home: const TestRunnerScreen(),
    );
  }
}

class TestRunnerScreen extends StatefulWidget {
  const TestRunnerScreen({super.key});

  @override
  State<TestRunnerScreen> createState() => TestRunnerScreenState();
}

class TestRunnerScreenState extends State<TestRunnerScreen> {
  late final List<TestCase> _tests;
  final Map<String, TestResult?> _results = {};
  bool _running = false;
  int _currentIndex = -1;
  int _passCount = 0;
  int _failCount = 0;
  int _totalDurationMs = 0;

  @override
  void initState() {
    super.initState();
    _tests = buildAllTests();
  }

  Future<void> runAll() async {
    setState(() {
      _running = true;
      _results.clear();
      _passCount = 0;
      _failCount = 0;
      _totalDurationMs = 0;
      _currentIndex = 0;
    });

    for (var i = 0; i < _tests.length; i++) {
      setState(() => _currentIndex = i);
      final result = await _tests[i].run();
      setState(() {
        _results[_tests[i].name] = result;
        if (result.passed) {
          _passCount++;
        } else {
          _failCount++;
        }
        _totalDurationMs += result.durationMs;
      });
    }

    setState(() {
      _running = false;
      _currentIndex = -1;
    });

    // Auto-save report after all tests complete
    await saveReport();
  }

  /// Generate a human-readable test report string.
  String generateReport() {
    final now = DateTime.now();
    final timestamp =
        '${now.year}-${now.month.toString().padLeft(2, '0')}-${now.day.toString().padLeft(2, '0')} '
        '${now.hour.toString().padLeft(2, '0')}:${now.minute.toString().padLeft(2, '0')}:${now.second.toString().padLeft(2, '0')}';

    final buf = StringBuffer();
    buf.writeln('╔══════════════════════════════════════════════════════════════╗');
    buf.writeln('║              excel_plus — Integration Test Report           ║');
    buf.writeln('╠══════════════════════════════════════════════════════════════╣');
    buf.writeln('║  Date     : $timestamp');
    buf.writeln('║  Platform : ${Platform.operatingSystem} ${Platform.operatingSystemVersion}');
    buf.writeln('║  Dart     : ${Platform.version.split(' ').first}');
    buf.writeln('╚══════════════════════════════════════════════════════════════╝');
    buf.writeln();

    // Column widths
    const numW = 4;
    const statusW = 6;
    const nameW = 25;
    const timeW = 10;
    const memW = 10;

    // Header
    buf.writeln('${'#'.padRight(numW)} ${'STATUS'.padRight(statusW)} ${'TEST NAME'.padRight(nameW)} ${'TIME'.padLeft(timeW)} ${'MEMORY'.padLeft(memW)}   MESSAGE');
    buf.writeln('${'─' * numW} ${'─' * statusW} ${'─' * nameW} ${'─' * timeW} ${'─' * memW}   ${'─' * 30}');

    var idx = 1;
    for (final test in _tests) {
      final result = _results[test.name];
      final num = idx.toString().padRight(numW);
      final status = result == null
          ? 'SKIP'.padRight(statusW)
          : (result.passed ? 'PASS'.padRight(statusW) : 'FAIL'.padRight(statusW));
      final name = test.name.padRight(nameW);
      final time = result != null
          ? '${result.durationMs}ms'.padLeft(timeW)
          : '-'.padLeft(timeW);
      final mem = result?.peakMemoryKB != null
          ? '${result!.peakMemoryKB}KB'.padLeft(memW)
          : '-'.padLeft(memW);
      final msg = result?.message ?? '';

      buf.writeln('$num $status $name $time $mem   $msg');
      idx++;
    }

    buf.writeln();
    buf.writeln('─' * 80);

    // Summary
    final total = _tests.length;
    final ran = _results.length;
    final skipped = total - ran;

    buf.writeln();
    buf.writeln('  SUMMARY');
    buf.writeln('  ├─ Total    : $total tests');
    buf.writeln('  ├─ Passed   : $_passCount');
    buf.writeln('  ├─ Failed   : $_failCount');
    if (skipped > 0) buf.writeln('  ├─ Skipped  : $skipped');
    buf.writeln('  ├─ Duration : ${_totalDurationMs}ms (${(_totalDurationMs / 1000).toStringAsFixed(1)}s)');
    buf.writeln('  └─ Result   : ${_failCount == 0 && ran == total ? '✅ ALL PASSED' : '❌ FAILURES DETECTED'}');
    buf.writeln();

    // Failures detail
    final failures = _tests.where((t) {
      final r = _results[t.name];
      return r != null && !r.passed;
    }).toList();

    if (failures.isNotEmpty) {
      buf.writeln('  FAILED TESTS:');
      for (final t in failures) {
        final r = _results[t.name]!;
        buf.writeln('    ✗ ${t.name} (${r.durationMs}ms)');
        buf.writeln('      ${r.message}');
      }
      buf.writeln();
    }

    return buf.toString();
  }

  /// Save the test report to the device's documents directory.
  Future<String?> saveReport() async {
    if (_results.isEmpty) return null;
    try {
      final dir = await getApplicationDocumentsDirectory();
      final reportDir = Directory('${dir.path}/excel_plus_test_reports');
      if (!reportDir.existsSync()) reportDir.createSync(recursive: true);

      final now = DateTime.now();
      final fileName =
          'test_report_${now.year}${now.month.toString().padLeft(2, '0')}${now.day.toString().padLeft(2, '0')}_'
          '${now.hour.toString().padLeft(2, '0')}${now.minute.toString().padLeft(2, '0')}${now.second.toString().padLeft(2, '0')}.txt';

      final file = File('${reportDir.path}/$fileName');
      final report = generateReport();
      await file.writeAsString(report);

      _lastReportPath = file.path;
      debugPrint('Report saved: ${file.path}');
      return file.path;
    } catch (e) {
      debugPrint('Failed to save report: $e');
      return null;
    }
  }

  String? _lastReportPath;
  String? get lastReportPath => _lastReportPath;

  Future<void> _runSingle(int index) async {
    setState(() {
      _running = true;
      _currentIndex = index;
    });

    final result = await _tests[index].run();
    setState(() {
      _results[_tests[index].name] = result;
      _running = false;
      _currentIndex = -1;
      _passCount = _results.values.where((r) => r != null && r.passed).length;
      _failCount =
          _results.values.where((r) => r != null && !r.passed).length;
      _totalDurationMs =
          _results.values.fold(0, (sum, r) => sum + (r?.durationMs ?? 0));
    });
  }

  /// Expose results for integration test assertions.
  Map<String, TestResult?> get results => _results;
  int get passCount => _passCount;
  int get failCount => _failCount;
  bool get isRunning => _running;

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('excel_plus Tests'),
        actions: [
          if (_results.isNotEmpty)
            Center(
              child: Padding(
                padding: const EdgeInsets.only(right: 16),
                child: Text(
                  '$_passCount✓  $_failCount✗  ${_totalDurationMs}ms',
                  style: const TextStyle(fontSize: 14),
                ),
              ),
            ),
        ],
      ),
      body: ListView.builder(
        key: const Key('test_list'),
        itemCount: _tests.length,
        itemBuilder: (context, index) {
          final test = _tests[index];
          final result = _results[test.name];
          final isCurrentlyRunning = _running && _currentIndex == index;

          return Card(
            key: Key('test_card_${test.name}'),
            margin: const EdgeInsets.symmetric(horizontal: 12, vertical: 4),
            child: ListTile(
              leading: _buildIcon(result, isCurrentlyRunning),
              title: Text(test.description),
              subtitle: result != null
                  ? Text(
                      result.message,
                      style: TextStyle(
                        color: result.passed
                            ? Colors.green[700]
                            : Colors.red[700],
                        fontSize: 12,
                      ),
                      maxLines: 2,
                      overflow: TextOverflow.ellipsis,
                    )
                  : Text(test.name,
                      style:
                          TextStyle(color: Colors.grey[500], fontSize: 12)),
              trailing: result != null
                  ? Text('${result.durationMs}ms',
                      style: const TextStyle(fontSize: 12))
                  : null,
              onTap: _running ? null : () => _runSingle(index),
            ),
          );
        },
      ),
      floatingActionButton: FloatingActionButton.extended(
        key: const Key('run_all_button'),
        onPressed: _running ? null : runAll,
        icon: _running
            ? const SizedBox(
                width: 20,
                height: 20,
                child: CircularProgressIndicator(strokeWidth: 2))
            : const Icon(Icons.play_arrow),
        label: Text(_running ? 'Running...' : 'Run All'),
      ),
    );
  }

  Widget _buildIcon(TestResult? result, bool isRunning) {
    if (isRunning) {
      return const SizedBox(
        width: 24,
        height: 24,
        child: CircularProgressIndicator(strokeWidth: 2),
      );
    }
    if (result == null) {
      return const Icon(Icons.circle_outlined, color: Colors.grey);
    }
    return result.passed
        ? const Icon(Icons.check_circle, color: Colors.green)
        : const Icon(Icons.error, color: Colors.red);
  }
}

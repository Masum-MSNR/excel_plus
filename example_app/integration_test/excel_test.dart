import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:integration_test/integration_test.dart';

import 'package:example_app/main.dart';
import 'package:example_app/tests/all_tests.dart';

void main() {
  IntegrationTestWidgetsFlutterBinding.ensureInitialized();

  group('excel_plus on-device tests', () {
    testWidgets('Run all tests and verify results', (tester) async {
      await tester.pumpWidget(const ExcelPlusTestApp());
      await tester.pumpAndSettle();

      // Tap the "Run All" button
      final runAllBtn = find.byKey(const Key('run_all_button'));
      expect(runAllBtn, findsOneWidget);
      await tester.tap(runAllBtn);

      // Wait for all tests to complete — poll until not running.
      // Each pump advances the frame and lets async test futures resolve.
      final state = tester.state<TestRunnerScreenState>(
          find.byType(TestRunnerScreen));

      // Give generous timeout for mobile (100K cell test can be slow)
      const maxWait = Duration(minutes: 5);
      final deadline = DateTime.now().add(maxWait);

      while (state.isRunning && DateTime.now().isBefore(deadline)) {
        await tester.pump(const Duration(milliseconds: 200));
      }
      await tester.pumpAndSettle();

      // Assert all tests ran
      final allTests = buildAllTests();
      expect(state.results.length, allTests.length,
          reason: 'Not all tests produced results');

      // Check each test individually for clear failure messages
      for (final test in allTests) {
        final result = state.results[test.name];
        expect(result, isNotNull, reason: '${test.name} has no result');
        expect(result!.passed, isTrue,
            reason: '${test.name} FAILED: ${result.message}');
      }

      // Summary
      expect(state.failCount, 0,
          reason:
              '${state.failCount} test(s) failed out of ${allTests.length}');

      debugPrint('==============================');
      debugPrint('ALL ${state.passCount} TESTS PASSED');
      debugPrint('Total duration: ${state.results.values.fold(0, (sum, r) => sum + (r?.durationMs ?? 0))}ms');
      debugPrint('==============================');
    });
  });
}

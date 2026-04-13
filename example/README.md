# excel_plus Example App

Flutter app for testing `excel_plus` on real Android devices and emulators.

## Features

13 tests covering the full `excel_plus` API:

| # | Test | What it verifies |
|---|------|-----------------|
| 1 | Create basic | New file, text cells, encode |
| 2 | Cell types | Text, int, double, bool, date, time, formula |
| 3 | Styles | Bold, italic, colors, borders via roundtrip |
| 4 | Multiple sheets | Create 3 sheets, verify after decode |
| 5 | Merge cells | Merge range, verify after roundtrip |
| 6 | Row/col operations | insertRow, data shift verification |
| 7 | Read existing | Open bundled .xlsx from assets |
| 8 | Roundtrip | 500 cells: create → encode → decode → compare |
| 9 | Column width/row height | Set and verify custom dimensions |
| 10 | Special characters | Unicode, emojis, XML entities, CJK |
| 11 | Large sheet 10K | 10,000 cells with timing |
| 12 | Large sheet 100K | 100,000 cells — mobile stress test |
| 13 | Save to disk | Write file to device storage |

## Manual Testing (UI)

```bash
cd example
flutter run
```

Tap **Run All** to execute all tests. Tap individual tests to re-run them.

## Automated Testing (Integration Tests on Emulator)

Uses Flutter's `integration_test` package — the official way to run
instrumented tests on real devices/emulators.

### Prerequisites

Android emulator running, or a physical device connected:
```bash
flutter devices
```

### Run all integration tests

```bash
cd example

# On a connected device/emulator:
flutter test integration_test/excel_test.dart

# Or specify a device:
flutter test integration_test/excel_test.dart -d emulator-5554
```

### Run with `flutter drive` (for CI / machine-readable output)

```bash
cd example
flutter drive \
  --driver=test_driver/integration_test.dart \
  --target=integration_test/excel_test.dart \
  -d emulator-5554
```

### CI Setup (GitHub Actions)

```yaml
jobs:
  integration-test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: subosito/flutter-action@v2
        with:
          channel: stable
      - name: Start emulator
        uses: reactivecircus/android-emulator-runner@v2
        with:
          api-level: 34
          script: |
            cd example
            flutter test integration_test/excel_test.dart
```

## Temp Files

Generated/temporary files go in `.tmp/` (gitignored).
- [Write your first Flutter app](https://docs.flutter.dev/get-started/codelab)
- [Flutter learning resources](https://docs.flutter.dev/reference/learning-resources)

For help getting started with Flutter development, view the
[online documentation](https://docs.flutter.dev/), which offers tutorials,
samples, guidance on mobile development, and a full API reference.

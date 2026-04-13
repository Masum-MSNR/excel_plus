# excel_plus

[![pub package](https://img.shields.io/pub/v/excel_plus.svg)](https://pub.dev/packages/excel_plus)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](https://opensource.org/licenses/MIT)

A fast, memory-efficient Dart & Flutter library for reading, creating, editing, and saving **Excel (.xlsx)** files. Drop-in replacement for the [`excel`](https://pub.dev/packages/excel) package with significantly better performance on large workbooks.

---

## Why excel_plus?

The original [`excel`](https://pub.dev/packages/excel) package works well for small files but struggles with large workbooks on mobile devices. `excel_plus` is a performance-optimized fork that keeps the same API while delivering:

- **7x faster** file open times on large workbooks
- **4x lower** peak memory usage
- **Lazy sheet loading** — only parses sheets you access
- **SAX-based streaming** — avoids building full DOM trees for cell data
- **Zero API changes** — switch your import and you're done

### Performance Comparison — Android Emulator (API 36)

Tested on Android emulator (x86_64), same device, same test data:

| Test | excel v5.0.0 | excel_plus v0.0.1 | Improvement |
|------|-------------|-------------------|-------------|
| **Create basic file** | ~160ms | **158ms** | — |
| **7 cell types roundtrip** | ~55ms | **53ms** | — |
| **Cell styling roundtrip** | ~20ms | **18ms** | — |
| **Multiple sheets (3)** | ~15ms | **12ms** | 20% faster |
| **Merge cells roundtrip** | ~22ms | **20ms** | — |
| **Row/column operations** | ~3ms | **2ms** | — |
| **Read existing .xlsx** | ~50ms | **47ms** | — |
| **500 cells roundtrip** | ~100ms | **91ms** | 9% faster |
| **Column width/row height** | ~18ms | **16ms** | — |
| **Special characters (unicode, emoji)** | ~20ms | **17ms** | — |
| **10,000 cells** (create→encode→decode) | ~250ms | **164ms** | **34% faster** |
| **100,000 cells** (create→encode→decode) | ~2,400ms | **768ms** | **68% faster** |
| **Save to disk** | ~65ms | **59ms** | — |
| | | | |
| **Total (13 tests)** | ~3,178ms | **1,425ms** | **55% faster** |

> The bigger the workbook, the bigger the improvement. On a 5M cell benchmark: **6.2s vs 46.3s open time, 2.7GB vs 11.5GB peak memory.**

---

## Features

- Read and write `.xlsx` files from bytes or streams
- Create new Excel workbooks from scratch
- Multiple sheets — create, copy, rename, delete, reorder
- Cell value types: `TextCellValue`, `IntCellValue`, `DoubleCellValue`, `BoolCellValue`, `DateCellValue`, `TimeCellValue`, `DateTimeCellValue`, `FormulaCellValue`
- Rich cell styling: fonts, colors, borders, alignment, rotation, number formats
- Merge and unmerge cells
- Insert, remove, and clear rows and columns
- Custom column widths and row heights with auto-fit
- Header and footer support
- Right-to-left (RTL) sheet support
- Find and replace across sheets
- Range selection and value extraction
- Works on all Dart platforms: VM, Web, mobile (Android & iOS)
- 100% API compatible with the `excel` package

---

## Getting Started

### Install

```yaml
dependencies:
  excel_plus: ^0.0.1
```

```bash
dart pub get
```

### Migrating from `excel`

Just change your import:

```dart
// Before
import 'package:excel/excel.dart';

// After
import 'package:excel_plus/excel_plus.dart';
```

No other code changes needed. All classes, methods, and enums are identical.

---

## Usage

### Read an Excel File

```dart
import 'dart:io';
import 'package:excel_plus/excel_plus.dart';

var file = File('path/to/file.xlsx');
var bytes = file.readAsBytesSync();
var excel = Excel.decodeBytes(bytes);

for (var table in excel.tables.keys) {
  print('Sheet: $table');
  for (var row in excel[table].rows) {
    print(row.map((cell) => cell?.value).toList());
  }
}
```

### Create a New Workbook

```dart
var excel = Excel.createExcel();
var sheet = excel['Sheet1'];

sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Name'));
sheet.updateCell(CellIndex.indexByString('B1'), TextCellValue('Age'));
sheet.updateCell(CellIndex.indexByString('A2'), TextCellValue('Alice'));
sheet.updateCell(CellIndex.indexByString('B2'), IntCellValue(30));

var bytes = excel.save();
File('output.xlsx').writeAsBytesSync(bytes!);
```

### Cell Styling

```dart
sheet.updateCell(
  CellIndex.indexByString('A1'),
  TextCellValue('Bold Red Header'),
  cellStyle: CellStyle(
    bold: true,
    fontSize: 14,
    fontColorHex: ExcelColor.fromHexString('#FF0000'),
    backgroundColorHex: ExcelColor.fromHexString('#FFFF00'),
    horizontalAlign: HorizontalAlign.Center,
    leftBorder: Border(borderStyle: BorderStyle.Thin),
    rightBorder: Border(borderStyle: BorderStyle.Thin),
    topBorder: Border(borderStyle: BorderStyle.Thin),
    bottomBorder: Border(borderStyle: BorderStyle.Thin),
  ),
);
```

### Merge Cells

```dart
sheet.merge(
  CellIndex.indexByString('A1'),
  CellIndex.indexByString('D1'),
  customValue: TextCellValue('Merged Header'),
);
```

### Row and Column Operations

```dart
sheet.insertRow(2);           // Insert empty row at index 2
sheet.removeRow(5);           // Remove row at index 5
sheet.insertColumn(1);        // Insert empty column at index 1
sheet.removeColumn(3);        // Remove column at index 3
sheet.appendRow([TextCellValue('a'), IntCellValue(1)]);
sheet.setColumnWidth(0, 25.0);
sheet.setRowHeight(0, 40.0);
```

### Multiple Sheets

```dart
var excel = Excel.createExcel();
excel['Sales'].updateCell(CellIndex.indexByString('A1'), TextCellValue('Revenue'));
excel['Inventory'].updateCell(CellIndex.indexByString('A1'), TextCellValue('Stock'));

excel.rename('Sales', 'Revenue');
excel.copy('Revenue', 'Revenue Backup');
excel.delete('Inventory');
excel.setDefaultSheet('Revenue');
```

### All Cell Value Types

```dart
sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Hello'));
sheet.updateCell(CellIndex.indexByString('A2'), IntCellValue(42));
sheet.updateCell(CellIndex.indexByString('A3'), DoubleCellValue(3.14));
sheet.updateCell(CellIndex.indexByString('A4'), BoolCellValue(true));
sheet.updateCell(CellIndex.indexByString('A5'), DateCellValue(year: 2026, month: 4, day: 14));
sheet.updateCell(CellIndex.indexByString('A6'), TimeCellValue(hour: 14, minute: 30, second: 0));
sheet.updateCell(CellIndex.indexByString('A7'), DateTimeCellValue(year: 2026, month: 4, day: 14, hour: 14, minute: 30));
sheet.updateCell(CellIndex.indexByString('A8'), FormulaCellValue('SUM(A2:A3)'));
```

### Flutter — Read from Assets and Save

```dart
import 'package:flutter/services.dart';
import 'package:path_provider/path_provider.dart';

// Read
final data = await rootBundle.load('assets/template.xlsx');
var excel = Excel.decodeBytes(data.buffer.asUint8List());

// Modify
excel['Sheet1'].updateCell(CellIndex.indexByString('A1'), TextCellValue('Updated'));

// Save
final dir = await getApplicationDocumentsDirectory();
File('${dir.path}/output.xlsx').writeAsBytesSync(excel.save()!);
```

---

## Architecture

```
lib/src/
  core/       → Excel class, configuration constants
  models/     → CellValue, CellStyle, CellIndex, enums, colors, borders
  sheet/      → Sheet class with row/column and merge mixins
  reader/     → SAX-based .xlsx parser with lazy sheet loading
  writer/     → .xlsx encoder with span correction
  utils/      → Archive, cell coordinate, and color utilities
  platform/   → Conditional imports for web vs native save
```

---

## License

MIT — see [LICENSE](LICENSE) for details.

---

## Credits

Based on the excellent [`excel`](https://pub.dev/packages/excel) package by [Kawal Jeet](https://github.com/justkawal). `excel_plus` is a performance-optimized fork focused on mobile and large-file workloads.

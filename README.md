<p align="center">
  <img src="https://raw.githubusercontent.com/Masum-MSNR/excel_plus/main/images/logo.png" alt="excel_plus" width="120"/>
</p>

<p align="center">
  <a href="https://pub.dev/packages/excel_plus"><img src="https://img.shields.io/pub/v/excel_plus.svg" alt="pub package"></a>
  <a href="https://opensource.org/licenses/MIT"><img src="https://img.shields.io/badge/License-MIT-blue.svg" alt="License: MIT"></a>
  <a href="https://dart.dev"><img src="https://img.shields.io/badge/Dart-3.11+-0175C2?logo=dart" alt="Dart"></a>
</p>

<p align="center">
A fast, memory-efficient Dart library for reading, creating, and editing Excel (.xlsx) files.<br/>
Drop-in replacement for the <a href="https://pub.dev/packages/excel">excel</a> package with significantly better performance on large workbooks.
</p>

## Features

- 📄 **Read & Write** `.xlsx` files from bytes
- 🆕 **Create** new Excel workbooks from scratch
- 📑 **Multiple Sheets** — create, copy, rename, delete, reorder
- 🔢 **All Cell Types** — Text, Int, Double, Bool, Date, Time, DateTime, Formula
- 🎨 **Cell Styling** — fonts, colors, borders, alignment, rotation, number formats
- 🔗 **Merge & Unmerge** cells with custom values
- ↕️ **Row & Column** — insert, remove, clear, resize
- 📐 **Custom Sizes** — column widths and row heights with auto-fit
- 🔄 **100% API Compatible** with the `excel` package
- 🌍 **Cross-Platform** — VM, Web, mobile (Android & iOS)

## Performance

Optimized with lazy sheet loading, SAX-based parsing, and smart memory management.

<p align="center">
  <img src="https://raw.githubusercontent.com/Masum-MSNR/excel_plus/main/images/benchmark.svg" alt="Performance benchmark" width="640"/>
</p>

## Installation

```yaml
dependencies:
  excel_plus: ^latest
```

### Migrating from `excel`

```dart
// Before
import 'package:excel/excel.dart';

// After
import 'package:excel_plus/excel_plus.dart';
```

No other code changes needed. All classes, methods, and enums are identical.

## Quick Start

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
sheet.insertRow(2);
sheet.removeRow(5);
sheet.insertColumn(1);
sheet.removeColumn(3);
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

final data = await rootBundle.load('assets/template.xlsx');
var excel = Excel.decodeBytes(data.buffer.asUint8List());

excel['Sheet1'].updateCell(CellIndex.indexByString('A1'), TextCellValue('Updated'));

final dir = await getApplicationDocumentsDirectory();
File('${dir.path}/output.xlsx').writeAsBytesSync(excel.save()!);
```

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

## License

MIT — see [LICENSE](LICENSE) for details.

---

<p align="center">
  Made with ❤️ by <a href="https://github.com/Masum-MSNR">Masum</a>
</p>

# excel_plus

A fast, memory-efficient Dart package for reading, creating, editing, and saving Excel (`.xlsx`) files.

## Features

- Read `.xlsx` files from bytes
- Create new Excel workbooks from scratch
- Edit cell values, styles, and formatting
- Support for multiple sheets
- Cell types: text, numbers, dates, times, booleans, formulas
- Cell styling: fonts, colors, borders, alignment
- Merge/unmerge cells
- Auto-fit column widths
- Header and footer support
- Works on all Dart platforms (VM, Web, mobile)

## Getting Started

```yaml
dependencies:
  excel_plus: ^0.0.1
```

## Usage

### Read an Excel file

```dart
import 'dart:io';
import 'package:excel_plus/excel_plus.dart';

var file = File('path/to/file.xlsx');
var bytes = file.readAsBytesSync();
var excel = Excel.decodeBytes(bytes);

for (var table in excel.tables.keys) {
  for (var row in excel.tables[table]!.rows) {
    print(row.map((cell) => cell?.value).toList());
  }
}
```

### Create and save

```dart
import 'package:excel_plus/excel_plus.dart';

var excel = Excel.createExcel();
var sheet = excel['Sheet1'];

sheet.cell(CellIndex.indexByString('A1')).value = TextCellValue('Hello');
sheet.cell(CellIndex.indexByString('B1')).value = IntCellValue(42);

var bytes = excel.save();
```

### Cell styling

```dart
var cell = sheet.cell(CellIndex.indexByString('A1'));
cell.value = TextCellValue('Styled');
cell.cellStyle = CellStyle(
  bold: true,
  fontSize: 14,
  fontColorHex: ExcelColor.fromHexString('#FF0000'),
);
```

## License

MIT

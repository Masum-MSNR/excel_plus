import 'dart:io';
import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

void main() {
  group('Excel basic operations', () {
    test('Create new excel file', () {
      var excel = Excel.createExcel();
      expect(excel.sheets, isNotEmpty);
    });

    test('Read and write cell value', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('test'),
      );
      var value = sheet.cell(CellIndex.indexByString('A1')).value;
      expect(value, isA<TextCellValue>());
      expect((value as TextCellValue).value.toString(), 'test');
    });

    test('Data class properties', () {
      var excel = Excel.createExcel();
      var sheet = excel['TestSheet'];
      sheet.updateCell(CellIndex.indexByString('C5'), TextCellValue('hello'));
      var data = sheet.cell(CellIndex.indexByString('C5'));
      expect(data.rowIndex, 4);
      expect(data.columnIndex, 2);
      expect(data.sheetName, 'TestSheet');
      expect(data.cellIndex, CellIndex.indexByString('C5'));
    });

    test('Data.setFormula', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(10));
      sheet.updateCell(CellIndex.indexByString('A2'), IntCellValue(20));
      var cell = sheet.cell(CellIndex.indexByString('A3'));
      cell.setFormula('SUM(A1:A2)');
      expect(cell.value, isA<FormulaCellValue>());
      expect((cell.value as FormulaCellValue).formula, 'SUM(A1:A2)');
    });

    test('CellIndex factories', () {
      var ci1 = CellIndex.indexByString('B3');
      expect(ci1.columnIndex, 1);
      expect(ci1.rowIndex, 2);
      expect(ci1.cellId, 'B3');

      var ci2 = CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: 0);
      expect(ci2.cellId, 'D1');
    });
  });

  group('CellValue roundtrip', () {
    test('Text, int, double, bool, formula cells roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('hello'));
      sheet.updateCell(CellIndex.indexByString('B1'), IntCellValue(42));
      sheet.updateCell(CellIndex.indexByString('C1'), DoubleCellValue(3.14));
      sheet.updateCell(CellIndex.indexByString('D1'), BoolCellValue(true));
      sheet.updateCell(CellIndex.indexByString('E1'), BoolCellValue(false));
      sheet.updateCell(
          CellIndex.indexByString('F1'), FormulaCellValue('SUM(B1,C1)'));
      sheet.updateCell(CellIndex.indexByString('A2'), TextCellValue('world'));
      sheet.updateCell(CellIndex.indexByString('B2'), IntCellValue(-100));
      sheet.updateCell(CellIndex.indexByString('C2'), DoubleCellValue(0.0));

      var bytes = excel.encode();
      expect(bytes, isNotNull);

      var decoded = Excel.decodeBytes(bytes!);
      var s = decoded['Sheet1'];

      expect(s.cell(CellIndex.indexByString('A1')).value, isA<TextCellValue>());
      expect(
          (s.cell(CellIndex.indexByString('A1')).value as TextCellValue)
              .value
              .toString(),
          'hello');

      expect(s.cell(CellIndex.indexByString('B1')).value, isA<IntCellValue>());
      expect(
          (s.cell(CellIndex.indexByString('B1')).value as IntCellValue).value,
          42);

      expect(
          s.cell(CellIndex.indexByString('C1')).value, isA<DoubleCellValue>());
      expect(
          (s.cell(CellIndex.indexByString('C1')).value as DoubleCellValue)
              .value,
          closeTo(3.14, 0.001));

      expect(
          s.cell(CellIndex.indexByString('D1')).value, isA<BoolCellValue>());
      expect(
          (s.cell(CellIndex.indexByString('D1')).value as BoolCellValue).value,
          true);

      expect(
          s.cell(CellIndex.indexByString('E1')).value, isA<BoolCellValue>());
      expect(
          (s.cell(CellIndex.indexByString('E1')).value as BoolCellValue).value,
          false);

      expect(s.cell(CellIndex.indexByString('F1')).value,
          isA<FormulaCellValue>());
      expect(
          (s.cell(CellIndex.indexByString('F1')).value as FormulaCellValue)
              .formula,
          'SUM(B1,C1)');

      expect(s.cell(CellIndex.indexByString('A2')).value, isA<TextCellValue>());
      expect(
          (s.cell(CellIndex.indexByString('A2')).value as TextCellValue)
              .value
              .toString(),
          'world');

      expect(s.cell(CellIndex.indexByString('B2')).value, isA<IntCellValue>());
      expect(
          (s.cell(CellIndex.indexByString('B2')).value as IntCellValue).value,
          -100);
    });

    test('DateCellValue roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        DateCellValue(year: 2024, month: 6, day: 15),
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var val = decoded['Sheet1'].cell(CellIndex.indexByString('A1')).value;
      expect(val, isA<DateCellValue>());
      var d = val as DateCellValue;
      expect(d.year, 2024);
      expect(d.month, 6);
      expect(d.day, 15);
    });

    test('TimeCellValue roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TimeCellValue(hour: 14, minute: 30, second: 45),
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var val = decoded['Sheet1'].cell(CellIndex.indexByString('A1')).value;
      expect(val, isA<TimeCellValue>());
      var t = val as TimeCellValue;
      expect(t.hour, 14);
      expect(t.minute, 30);
      expect(t.second, 45);
    });

    test('DateTimeCellValue roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        DateTimeCellValue(
            year: 2025, month: 12, day: 25, hour: 10, minute: 30, second: 15),
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var val = decoded['Sheet1'].cell(CellIndex.indexByString('A1')).value;
      expect(val, isA<DateTimeCellValue>());
      var dt = val as DateTimeCellValue;
      expect(dt.year, 2025);
      expect(dt.month, 12);
      expect(dt.day, 25);
      expect(dt.hour, 10);
      expect(dt.minute, 30);
      expect(dt.second, 15);
    });

    test('Null cell value roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('data'));

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var val = decoded['Sheet1'].cell(CellIndex.indexByString('B1')).value;
      expect(val, isNull);
    });

    test('Special characters in text roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(
          CellIndex.indexByString('A1'), TextCellValue('a & b < c > d "e"'));
      sheet.updateCell(
          CellIndex.indexByString('A2'), TextCellValue("it's a test"));
      sheet.updateCell(CellIndex.indexByString('A3'),
          TextCellValue('Unicode: \u00e9\u00f1\u00fc \u4e16\u754c'));
      sheet.updateCell(
          CellIndex.indexByString('A4'), TextCellValue('Emoji: \u{1F600}'));

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var s = decoded['Sheet1'];

      expect(
          (s.cell(CellIndex.indexByString('A1')).value as TextCellValue)
              .value
              .toString(),
          'a & b < c > d "e"');
      expect(
          (s.cell(CellIndex.indexByString('A2')).value as TextCellValue)
              .value
              .toString(),
          "it's a test");
      expect(
          (s.cell(CellIndex.indexByString('A3')).value as TextCellValue)
              .value
              .toString(),
          'Unicode: \u00e9\u00f1\u00fc \u4e16\u754c');
      expect(
          (s.cell(CellIndex.indexByString('A4')).value as TextCellValue)
              .value
              .toString(),
          'Emoji: \u{1F600}');
    });

    test('Many rows/columns roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      for (var r = 0; r < 100; r++) {
        for (var c = 0; c < 20; c++) {
          sheet.updateCell(
            CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
            TextCellValue('R${r}C$c'),
          );
        }
      }

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var s = decoded['Sheet1'];

      for (var r = 0; r < 100; r++) {
        for (var c = 0; c < 20; c++) {
          var val = s
              .cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r))
              .value;
          expect(val, isA<TextCellValue>(),
              reason: 'Cell R${r}C$c should be TextCellValue');
          expect((val as TextCellValue).value.toString(), 'R${r}C$c',
              reason: 'Cell R${r}C$c value mismatch');
        }
      }
    });
  });

  group('Read existing XLSX files', () {
    test('Read example.xlsx', () {
      var file = './test/test_resources/example.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      expect(excel.tables['Sheet1']!.maxColumns, equals(3));
      expect(excel.tables['Sheet1']!.rows[1][1]!.value.toString(),
          equals('Washington'));
    });

    test('Read data types from MS Excel 365', () {
      var file = './test/test_resources/dataTypesUsingMsExcel365Desktop.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      expect(excel.tables['Tabelle1']!.rows[2][1]?.value,
          equals(TextCellValue('Some text')));
      expect(excel.tables['Tabelle1']?.rows[3][1]?.value,
          equals(IntCellValue(42)));
      expect(excel.tables['Tabelle1']?.rows[4][1]?.value,
          equals(DoubleCellValue(12.3)));
      expect(excel.tables['Tabelle1']?.rows[7][1]?.value,
          equals(BoolCellValue(true)));
      expect(excel.tables['Tabelle1']?.rows[8][1]?.value,
          equals(BoolCellValue(false)));
    });

    test('Read + encode + decode roundtrip on existing file', () {
      var file = './test/test_resources/example.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);

      excel['Sheet1'].updateCell(
        CellIndex.indexByString('D1'),
        TextCellValue('NewColumn'),
      );

      var encoded = excel.encode();
      expect(encoded, isNotNull);
      var decoded = Excel.decodeBytes(encoded!);

      expect(decoded.tables['Sheet1']!.rows[1][1]!.value.toString(),
          equals('Washington'));
      expect(decoded['Sheet1'].cell(CellIndex.indexByString('D1')).value,
          isA<TextCellValue>());
      expect(
          (decoded['Sheet1'].cell(CellIndex.indexByString('D1')).value
                  as TextCellValue)
              .value
              .toString(),
          'NewColumn');
    });

    test('Read spannedItemExample.xlsx', () {
      var file = './test/test_resources/spannedItemExample.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      var sheet = excel.tables.values.first;
      expect(sheet.spannedItems, isNotEmpty);
    });

    test('Read borders.xlsx', () {
      var file = './test/test_resources/borders.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      expect(excel.tables, isNotEmpty);
    });

    test('Read columnWidthRowHeight.xlsx', () {
      var file = './test/test_resources/columnWidthRowHeight.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      expect(excel.tables, isNotEmpty);
    });

    test('Read data types from Google Spreadsheet', () {
      var file = './test/test_resources/dataTypesUsingGoogleSpreadsheet.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      expect(excel.tables, isNotEmpty);
    });

    test('Read data types from LibreOffice', () {
      var file = './test/test_resources/dataTypesUsingLibreoffice.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      expect(excel.tables, isNotEmpty);
    });
  });
}

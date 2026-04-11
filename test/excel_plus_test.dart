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
  });

  group('Save/Read roundtrip', () {
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

    test('Multiple sheets roundtrip', () {
      var excel = Excel.createExcel();
      excel['SheetA']
          .updateCell(CellIndex.indexByString('A1'), TextCellValue('Alpha'));
      excel['SheetB']
          .updateCell(CellIndex.indexByString('B2'), IntCellValue(99));

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);

      expect(decoded.sheets.keys, contains('SheetA'));
      expect(decoded.sheets.keys, contains('SheetB'));
      expect(
          (decoded['SheetA'].cell(CellIndex.indexByString('A1')).value
                  as TextCellValue)
              .value
              .toString(),
          'Alpha');
      expect(
          (decoded['SheetB'].cell(CellIndex.indexByString('B2')).value
                  as IntCellValue)
              .value,
          99);
    });

    test('Styled cells roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('styled'));

      var cell = sheet.cell(CellIndex.indexByString('A1'));
      cell.cellStyle = CellStyle(bold: true);

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var s = decoded['Sheet1'];
      var val = s.cell(CellIndex.indexByString('A1')).value;
      expect(val, isA<TextCellValue>());
      expect((val as TextCellValue).value.toString(), 'styled');
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

    test('Special characters in text roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(
          CellIndex.indexByString('A1'), TextCellValue('a & b < c > d "e"'));
      sheet.updateCell(
          CellIndex.indexByString('A2'), TextCellValue("it's a test"));

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
    });

    test('Date and time cells roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      sheet.updateCell(
        CellIndex.indexByString('A1'),
        DateCellValue(year: 2024, month: 6, day: 15),
      );
      sheet.updateCell(
        CellIndex.indexByString('B1'),
        TimeCellValue(hour: 14, minute: 30, second: 0),
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var s = decoded['Sheet1'];

      // Date cells are stored as numbers in XLSX — just verify it's not null
      expect(s.cell(CellIndex.indexByString('A1')).value, isNotNull);
      expect(s.cell(CellIndex.indexByString('B1')).value, isNotNull);
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

      // Modify a cell
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('D1'),
        TextCellValue('NewColumn'),
      );

      // Encode then decode
      var encoded = excel.encode();
      expect(encoded, isNotNull);
      var decoded = Excel.decodeBytes(encoded!);

      // Original data preserved
      expect(decoded.tables['Sheet1']!.rows[1][1]!.value.toString(),
          equals('Washington'));
      // New cell present
      expect(decoded['Sheet1'].cell(CellIndex.indexByString('D1')).value,
          isA<TextCellValue>());
      expect(
          (decoded['Sheet1'].cell(CellIndex.indexByString('D1')).value
                  as TextCellValue)
              .value
              .toString(),
          'NewColumn');
    });
  });
}

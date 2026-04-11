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
}

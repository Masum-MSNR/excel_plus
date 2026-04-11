import 'package:excel_plus/excel_plus.dart';

void main() {
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];
  sheet.updateCell(
    CellIndex.indexByString('A1'),
    TextCellValue('Hello from excel_plus'),
  );
  print('Sheets: ${excel.sheets}');
}

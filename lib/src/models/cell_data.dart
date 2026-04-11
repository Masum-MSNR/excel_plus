part of '../../excel_plus.dart';

class Data {
  CellStyle? _cellStyle;
  CellValue? _value;
  Sheet _sheet;
  int _rowIndex;
  int _columnIndex;

  Data._clone(Sheet sheet, Data dataObject)
      : this._(
          sheet,
          dataObject._rowIndex,
          dataObject.columnIndex,
          value: dataObject._value,
          cellStyleVal: dataObject._cellStyle,
        );

  Data._(
    Sheet sheet,
    int row,
    int column, {
    CellValue? value,
    NumFormat? numberFormat,
    CellStyle? cellStyleVal,
    bool isFormulaVal = false,
  })  : _sheet = sheet,
        _value = value,
        _cellStyle = cellStyleVal,
        _rowIndex = row,
        _columnIndex = column;

  static Data newData(Sheet sheet, int row, int column) {
    return Data._(sheet, row, column);
  }

  /// returns the row Index
  int get rowIndex => _rowIndex;

  /// returns the column Index
  int get columnIndex => _columnIndex;

  /// returns the sheet-name
  String get sheetName => _sheet.sheetName;

  /// returns the string based cellId as A1, A2 or Z5
  CellIndex get cellIndex {
    return CellIndex.indexByColumnRow(
        columnIndex: _columnIndex, rowIndex: _rowIndex);
  }

  /// Helps to set the formula
  void setFormula(String formula) {
    _sheet.updateCell(cellIndex, FormulaCellValue(formula));
  }

  set value(CellValue? val) {
    _sheet.updateCell(cellIndex, val);
  }

  /// returns the value stored in this cell;
  ///
  /// It will return `null` if no value is stored in this cell.
  CellValue? get value => _value;

  /// returns the user-defined CellStyle
  ///
  /// if `no` cellStyle is set then it returns `null`
  CellStyle? get cellStyle {
    return _cellStyle;
  }

  /// sets the user defined CellStyle in this current cell
  set cellStyle(CellStyle? style) {
    _sheet._excel._styleChanges = true;
    _cellStyle = style;
  }

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is Data &&
          _rowIndex == other._rowIndex &&
          _columnIndex == other._columnIndex &&
          _value == other._value &&
          _cellStyle == other._cellStyle;

  @override
  int get hashCode => Object.hash(_rowIndex, _columnIndex, _value, _cellStyle);
}

part of '../../excel_plus.dart';

/// Base class containing ExcelWriter fields and cell/row utility methods.
///
/// Not meant to be used directly. Use [ExcelWriter] instead.
abstract class _WriterBase {
  final Excel _excel;
  final Map<String, ArchiveFile> _archiveFiles = {};
  final Map<CellStyle, int> _innerCellStyle = {};
  final Parser parser;

  _WriterBase(this._excel, this.parser);

  void _addNewColumn(XmlElement columns, int min, int max, double width) {
    columns.children.add(XmlElement(XmlName('col'), [
      XmlAttribute(XmlName('min'), (min + 1).toString()),
      XmlAttribute(XmlName('max'), (max + 1).toString()),
      XmlAttribute(XmlName('width'), width.toStringAsFixed(2)),
      XmlAttribute(XmlName('bestFit'), "1"),
      XmlAttribute(XmlName('customWidth'), "1"),
    ], []));
  }

  double _calcAutoFitColumnWidth(Sheet sheet, int column) {
    var maxNumOfCharacters = 0;
    sheet._sheetData.forEach((key, value) {
      if (value.containsKey(column) &&
          value[column]!.value is! FormulaCellValue) {
        maxNumOfCharacters =
            max(value[column]!.value.toString().length, maxNumOfCharacters);
      }
    });

    return ((maxNumOfCharacters * 7.0 + 9.0) / 7.0 * 256).truncate() / 256;
  }

  // Manage value's type
  XmlElement _createCell(String sheet, int columnIndex, int rowIndex,
      CellValue? value, NumFormat? numberFormat) {
    SharedString? sharedString;
    if (value is TextCellValue) {
      sharedString = _excel._sharedStrings.tryFind(value.toString());
      if (sharedString != null) {
        _excel._sharedStrings.add(sharedString, value.toString());
      } else {
        sharedString = _excel._sharedStrings.addFromString(value.toString());
      }
    }

    String rC = getCellId(columnIndex, rowIndex);

    var attributes = <XmlAttribute>[
      XmlAttribute(XmlName('r'), rC),
      if (value is TextCellValue) XmlAttribute(XmlName('t'), 's'),
      if (value is BoolCellValue) XmlAttribute(XmlName('t'), 'b'),
    ];

    final cellStyle =
        _excel._sheetMap[sheet]?._sheetData[rowIndex]?[columnIndex]?.cellStyle;

    if (_excel._styleChanges && cellStyle != null) {
      int upperLevelPos = _excel._cellStyleList.indexOf(cellStyle);
      if (upperLevelPos == -1) {
        int lowerLevelPos = _innerCellStyle[cellStyle] ?? -1;
        if (lowerLevelPos != -1) {
          upperLevelPos = lowerLevelPos + _excel._cellStyleList.length;
        } else {
          upperLevelPos = 0;
        }
      }
      attributes.insert(
        1,
        XmlAttribute(XmlName('s'), '$upperLevelPos'),
      );
    } else if (_excel._cellStyleReferenced.containsKey(sheet) &&
        _excel._cellStyleReferenced[sheet]!.containsKey(rC)) {
      attributes.insert(
        1,
        XmlAttribute(
            XmlName('s'), '${_excel._cellStyleReferenced[sheet]![rC]}'),
      );
    }

    // TODO track & write the numFmts/numFmt to styles.xml if used
    final List<XmlElement> children;
    switch (value) {
      case null:
        children = [];
      case FormulaCellValue():
        children = [
          XmlElement(XmlName('f'), [], [XmlText(value.formula)]),
          XmlElement(XmlName('v'), [], [XmlText('')]),
        ];
      case IntCellValue():
        final String v = switch (numberFormat) {
          NumericNumFormat() => numberFormat.writeInt(value),
          _ => throw Exception(
              '$numberFormat does not work for ${value.runtimeType}'),
        };
        children = [
          XmlElement(XmlName('v'), [], [XmlText(v)]),
        ];
      case DoubleCellValue():
        final String v = switch (numberFormat) {
          NumericNumFormat() => numberFormat.writeDouble(value),
          _ => throw Exception(
              '$numberFormat does not work for ${value.runtimeType}'),
        };
        children = [
          XmlElement(XmlName('v'), [], [XmlText(v)]),
        ];
      case DateTimeCellValue():
        final String v = switch (numberFormat) {
          DateTimeNumFormat() => numberFormat.writeDateTime(value),
          _ => throw Exception(
              '$numberFormat does not work for ${value.runtimeType}'),
        };
        children = [
          XmlElement(XmlName('v'), [], [XmlText(v)]),
        ];
      case DateCellValue():
        final String v = switch (numberFormat) {
          DateTimeNumFormat() => numberFormat.writeDate(value),
          _ => throw Exception(
              '$numberFormat does not work for ${value.runtimeType}'),
        };
        children = [
          XmlElement(XmlName('v'), [], [XmlText(v)]),
        ];
      case TimeCellValue():
        final String v = switch (numberFormat) {
          TimeNumFormat() => numberFormat.writeTime(value),
          _ => throw Exception(
              '$numberFormat does not work for ${value.runtimeType}'),
        };
        children = [
          XmlElement(XmlName('v'), [], [XmlText(v)]),
        ];
      case TextCellValue():
        children = [
          XmlElement(XmlName('v'), [], [
            XmlText(_excel._sharedStrings.indexOf(sharedString!).toString())
          ]),
        ];
      case BoolCellValue():
        children = [
          XmlElement(XmlName('v'), [], [XmlText(value.value ? '1' : '0')]),
        ];
    }

    return XmlElement(XmlName('c'), attributes, children);
  }

  /// Create a new row in the sheet.
  XmlElement _createNewRow(XmlElement table, int rowIndex, double? height) {
    var row = XmlElement(XmlName('row'), [
      XmlAttribute(XmlName('r'), (rowIndex + 1).toString()),
      if (height != null)
        XmlAttribute(XmlName('ht'), height.toStringAsFixed(2)),
      if (height != null) XmlAttribute(XmlName('customHeight'), '1'),
    ], []);
    table.children.add(row);
    return row;
  }

  XmlElement _updateCell(String sheet, XmlElement row, int columnIndex,
      int rowIndex, CellValue? value, NumFormat? numberFormat) {
    var cell = _createCell(sheet, columnIndex, rowIndex, value, numberFormat);
    row.children.add(cell);
    return cell;
  }

  _BorderSet _createBorderSetFromCellStyle(CellStyle cellStyle) => _BorderSet(
        leftBorder: cellStyle.leftBorder,
        rightBorder: cellStyle.rightBorder,
        topBorder: cellStyle.topBorder,
        bottomBorder: cellStyle.bottomBorder,
        diagonalBorder: cellStyle.diagonalBorder,
        diagonalBorderUp: cellStyle.diagonalBorderUp,
        diagonalBorderDown: cellStyle.diagonalBorderDown,
      );
}

part of '../../excel_plus.dart';

class Parser extends _ParserBase with _ParserStylesMixin {
  Parser._(super.excel);

  void _startParsing() {
    _putContentXml();
    _parseRelations();
    _parseStyles(_excel._stylesTarget);
    _parseSharedStrings();
    _parseContent();
  }

  @override
  void _parseContent({bool run = true}) {
    var workbook = _excel._archive.findFile('xl/workbook.xml');
    if (workbook == null) {
      _damagedExcel();
    }
    workbook!.decompress();
    var document = XmlDocument.parse(utf8.decode(workbook.content));
    _excel._xmlFiles["xl/workbook.xml"] = document;

    document.findAllElements('sheet').forEach((node) {
      var name = node.getAttribute('name');
      var rid = node.getAttribute('r:id');
      if (name != null) {
        // Create empty Sheet object so sheet names are visible immediately
        if (_excel._sheetMap[name] == null) {
          _excel._sheetMap[name] = Sheet._(_excel, name);
        }
        // Store node for deferred parsing
        _excel._pendingSheetNodes[name] = node;
      }
      if (!run && rid != null && !_rId.contains(rid)) {
        _rId.add(rid);
      }
    });
  }

  /// Parses a single sheet on demand. Called from [Excel._availSheet].
  void _ensureSheetParsed(String sheetName) {
    final node = _excel._pendingSheetNodes.remove(sheetName);
    if (node == null) return;
    _parseTable(node);
    _parseMergedCellsForSheet(sheetName);
  }

  /// Parses all remaining unparsed sheets.
  void _ensureAllSheetsParsed() {
    if (_excel._pendingSheetNodes.isEmpty) return;
    for (final name in _excel._pendingSheetNodes.keys.toList()) {
      _ensureSheetParsed(name);
    }
  }

  /// Parses merged cells for a single sheet.
  void _parseMergedCellsForSheet(String sheetName) {
    final node = _excel._sheets[sheetName];
    if (node == null) return;
    _excel._availSheet(sheetName);
    XmlElement sheetDataNode = node as XmlElement;
    final sheet = _excel._sheetMap[sheetName]!;

    final worksheetNode = sheetDataNode.parent;
    worksheetNode!.findAllElements('mergeCell').forEach((element) {
      String? ref = element.getAttribute('ref');
      if (ref != null && ref.contains(':') && ref.split(':').length == 2) {
        if (!sheet._spannedItems.contains(ref)) {
          sheet._spannedItems.add(ref);
        }

        String startCell = ref.split(':')[0], endCell = ref.split(':')[1];

        CellIndex startIndex = CellIndex.indexByString(startCell),
            endIndex = CellIndex.indexByString(endCell);
        _Span spanObj = _Span.fromCellIndex(
          start: startIndex,
          end: endIndex,
        );
        if (!sheet._spanList.contains(spanObj)) {
          sheet._spanList.add(spanObj);
          _deleteAllButTopLeftCellsOfSpanObj(spanObj, sheet);
        }
        _excel._mergeChangeLookup = sheetName;
      }
    });
  }

  /// Deletes all cells within the span of the given [_Span] object
  /// except for the top-left cell.
  ///
  /// This method is used internally by [_parseMergedCells] to remove
  /// cells within merged cell regions.
  ///
  /// Parameters:
  ///   - [spanObj]: The span object representing the merged cell region.
  ///   - [sheet]: The sheet object from which cells are to be removed.
  void _deleteAllButTopLeftCellsOfSpanObj(_Span spanObj, Sheet sheet) {
    final columnSpanStart = spanObj.columnSpanStart;
    final columnSpanEnd = spanObj.columnSpanEnd;
    final rowSpanStart = spanObj.rowSpanStart;
    final rowSpanEnd = spanObj.rowSpanEnd;

    for (var columnI = columnSpanStart; columnI <= columnSpanEnd; columnI++) {
      for (var rowI = rowSpanStart; rowI <= rowSpanEnd; rowI++) {
        bool isTopLeftCellThatShouldNotBeDeleted =
            columnI == columnSpanStart && rowI == rowSpanStart;

        if (isTopLeftCellThatShouldNotBeDeleted) {
          continue;
        }
        sheet._removeCell(rowI, columnI);
      }
    }
  }

  void _parseTable(XmlElement node) {
    var name = node.getAttribute('name')!;
    var target = _worksheetTargets[node.getAttribute('r:id')];

    if (_excel._sheetMap[name] == null) {
      _excel._sheetMap[name] = Sheet._(_excel, name);
    }

    Sheet sheetObject = _excel._sheetMap[name]!;

    var file = _excel._archive.findFile('xl/$target');
    file!.decompress();

    var content = XmlDocument.parse(utf8.decode(file.content));
    var worksheet = content.findElements('worksheet').first;

    ///
    /// check for right to left view
    ///
    var sheetView = worksheet.findAllElements('sheetView').toList();
    if (sheetView.isNotEmpty) {
      var sheetViewNode = sheetView.first;
      var rtl = sheetViewNode.getAttribute('rightToLeft');
      sheetObject.isRTL = rtl != null && rtl == '1';
    }
    var sheet = worksheet.findElements('sheetData').first;

    _findRows(sheet).forEach((child) {
      _parseRow(child, sheetObject, name);
    });

    _parseHeaderFooter(worksheet, sheetObject);
    _parseColWidthsRowHeights(worksheet, sheetObject);

    _excel._sheets[name] = sheet;

    _excel._xmlFiles['xl/$target'] = content;
    _excel._xmlSheetId[name] = 'xl/$target';

    _normalizeTable(sheetObject);
  }

  void _parseRow(XmlElement node, Sheet sheetObject, String name) {
    var rowIndex = (_getRowNumber(node) ?? -1) - 1;
    if (rowIndex < 0) {
      return;
    }

    _findCells(node).forEach((child) {
      _parseCell(child, sheetObject, rowIndex, name);
    });
  }

  void _parseCell(
      XmlElement node, Sheet sheetObject, int rowIndex, String name) {
    int? columnIndex = _getCellNumber(node);
    if (columnIndex == null) {
      return;
    }

    var s1 = node.getAttribute('s');
    int s = 0;
    if (s1 != null) {
      try {
        s = int.parse(s1.toString());
      } catch (_) {}

      String rC = node.getAttribute('r').toString();

      if (_excel._cellStyleReferenced[name] == null) {
        _excel._cellStyleReferenced[name] = {rC: s};
      } else {
        _excel._cellStyleReferenced[name]![rC] = s;
      }
    }

    CellValue? value;
    String? type = node.getAttribute('t');

    switch (type) {
      // sharedString
      case 's':
        final sharedString = _excel._sharedStrings
            .value(int.parse(_parseValue(node.findElements('v').first)));
        value = TextCellValue.span(sharedString!.textSpan);
        break;
      // boolean
      case 'b':
        value = BoolCellValue(_parseValue(node.findElements('v').first) == '1');
        break;
      // error
      case 'e':
      // formula
      case 'str':
        value = FormulaCellValue(_parseValue(node.findElements('v').first));
        break;
      // inline string
      case 'inlineStr':
        // <c r='B2' t='inlineStr'>
        // <is><t>Dartonico</t></is>
        // </c>
        value = TextCellValue(_parseValue(node.findAllElements('t').first));
        break;
      // number
      case 'n':
      default:
        var formulaNode = node.findElements('f');
        if (formulaNode.isNotEmpty) {
          value = FormulaCellValue(_parseValue(formulaNode.first).toString());
        } else {
          final vNode = node.findElements('v').firstOrNull;
          if (vNode == null) {
            value = null;
          } else if (s1 != null) {
            final v = _parseValue(vNode);
            var numFmtId = _excel._numFmtIds[s];
            final numFormat = _excel._numFormats.getByNumFmtId(numFmtId);
            if (numFormat == null) {
              assert(
                  false, 'found no number format spec for numFmtId $numFmtId');
              value = NumFormat.defaultNumeric.read(v);
            } else {
              value = numFormat.read(v);
            }
          } else {
            final v = _parseValue(vNode);
            value = NumFormat.defaultNumeric.read(v);
          }
        }
    }

    sheetObject.updateCell(
      CellIndex.indexByColumnRow(columnIndex: columnIndex, rowIndex: rowIndex),
      value,
      cellStyle: _excel._cellStyleList[s],
    );
  }

  static String _parseValue(XmlElement node) {
    var buffer = StringBuffer();

    for (var child in node.children) {
      if (child is XmlText) {
        buffer.write(_normalizeNewLine(child.value));
      }
    }

    return buffer.toString();
  }

  ///Uses the [newSheet] as the name of the sheet and also adds it to the [ xl/worksheets/ ] directory
  ///
  ///Creates the sheet with name `newSheet` as file output and then adds it to the archive directory.
  ///
  ///
  void _createSheet(String newSheet) {
    /*
    List<XmlNode> list = _excel._xmlFiles['xl/workbook.xml']
        .findAllElements('sheets')
        .first
        .children;
    if (list.isEmpty) {
      throw ArgumentError('');
    } */

    int sheetId0 = -1;
    List<int> sheetIdList = <int>[];

    _excel._xmlFiles['xl/workbook.xml']
        ?.findAllElements('sheet')
        .forEach((sheetIdNode) {
      var sheetId = sheetIdNode.getAttribute('sheetId');
      if (sheetId != null) {
        int t = int.parse(sheetId.toString());
        if (!sheetIdList.contains(t)) {
          sheetIdList.add(t);
        }
      } else {
        _damagedExcel(text: 'Corrupted Sheet Indexing');
      }
    });

    sheetIdList.sort();

    for (int i = 0; i < sheetIdList.length; i++) {
      if ((i + 1) != sheetIdList[i]) {
        sheetId0 = i + 1;
        break;
      }
    }
    if (sheetId0 == -1) {
      if (sheetIdList.isEmpty) {
        sheetId0 = 1;
      } else {
        sheetId0 = sheetIdList.length + 1;
      }
    }

    int sheetNumber = sheetId0;
    int ridNumber = _getAvailableRid();

    _excel._xmlFiles['xl/_rels/workbook.xml.rels']
        ?.findAllElements('Relationships')
        .first
        .children
        .add(XmlElement(XmlName('Relationship'), <XmlAttribute>[
          XmlAttribute(XmlName('Id'), 'rId$ridNumber'),
          XmlAttribute(XmlName('Type'), '$_relationships/worksheet'),
          XmlAttribute(XmlName('Target'), 'worksheets/sheet$sheetNumber.xml'),
        ]));

    if (!_rId.contains('rId$ridNumber')) {
      _rId.add('rId$ridNumber');
    }

    _excel._xmlFiles['xl/workbook.xml']
        ?.findAllElements('sheets')
        .first
        .children
        .add(XmlElement(
          XmlName('sheet'),
          <XmlAttribute>[
            XmlAttribute(XmlName('state'), 'visible'),
            XmlAttribute(XmlName('name'), newSheet),
            XmlAttribute(XmlName('sheetId'), '$sheetNumber'),
            XmlAttribute(XmlName('r:id'), 'rId$ridNumber')
          ],
        ));

    _worksheetTargets['rId$ridNumber'] = 'worksheets/sheet$sheetNumber.xml';

    var content = utf8.encode(
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac xr xr2 xr3\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\"> <dimension ref=\"A1\"/> <sheetViews> <sheetView workbookViewId=\"0\"/> </sheetViews> <sheetData/> <pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/> </worksheet>");

    _excel._archive.addFile(ArchiveFile(
        'xl/worksheets/sheet$sheetNumber.xml', content.length, content));
    var newSheet0 =
        _excel._archive.findFile('xl/worksheets/sheet$sheetNumber.xml');

    newSheet0!.decompress();
    var document = XmlDocument.parse(utf8.decode(newSheet0.content));
    _excel._xmlFiles['xl/worksheets/sheet$sheetNumber.xml'] = document;
    _excel._xmlSheetId[newSheet] = 'xl/worksheets/sheet$sheetNumber.xml';

    _excel._xmlFiles['[Content_Types].xml']
        ?.findAllElements('Types')
        .first
        .children
        .add(XmlElement(
          XmlName('Override'),
          <XmlAttribute>[
            XmlAttribute(XmlName('ContentType'),
                'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
            XmlAttribute(
                XmlName('PartName'), '/xl/worksheets/sheet$sheetNumber.xml'),
          ],
        ));
    if (_excel._xmlFiles['xl/workbook.xml'] != null) {
      _parseTable(
          _excel._xmlFiles['xl/workbook.xml']!.findAllElements('sheet').last);
    }
  }

  void _parseHeaderFooter(XmlElement worksheet, Sheet sheetObject) {
    final results = worksheet.findAllElements("headerFooter");
    if (results.isEmpty) return;

    final headerFooterElement = results.first;

    sheetObject.headerFooter = HeaderFooter.fromXmlElement(headerFooterElement);
  }

  void _parseColWidthsRowHeights(XmlElement worksheet, Sheet sheetObject) {
    /* parse default column width and default row height
      example XML content
      <sheetFormatPr baseColWidth="10" defaultColWidth="26.33203125" defaultRowHeight="13" x14ac:dyDescent="0.15" />
    */
    Iterable<XmlElement> results;
    results = worksheet.findAllElements("sheetFormatPr");
    if (results.isNotEmpty) {
      for (var element in results) {
        double? defaultColWidth;
        double? defaultRowHeight;
        // default column width
        String? widthAttribute = element.getAttribute("defaultColWidth");
        if (widthAttribute != null) {
          defaultColWidth = double.tryParse(widthAttribute);
        }
        // default row height
        String? rowHeightAttribute = element.getAttribute("defaultRowHeight");
        if (rowHeightAttribute != null) {
          defaultRowHeight = double.tryParse(rowHeightAttribute);
        }

        // both values valid ?
        if (defaultColWidth != null && defaultRowHeight != null) {
          sheetObject._defaultColumnWidth = defaultColWidth;
          sheetObject._defaultRowHeight = defaultRowHeight;
        }
      }
    }

    /* parse custom column height
      example XML content
      <col min="2" max="2" width="71.83203125" customWidth="1"/>, 
      <col min="4" max="4" width="26.5" customWidth="1"/>, 
      <col min="6" max="6" width="31.33203125" customWidth="1"/>
    */
    results = worksheet.findAllElements("col");
    if (results.isNotEmpty) {
      for (var element in results) {
        String? colAttribute =
            element.getAttribute("min"); // i think min refers to the column
        String? widthAttribute = element.getAttribute("width");
        if (colAttribute != null && widthAttribute != null) {
          int? col = int.tryParse(colAttribute);
          double? width = double.tryParse(widthAttribute);
          if (col != null && width != null) {
            col -= 1; // first col in _columnWidths is index 0
            if (col >= 0) {
              sheetObject._columnWidths[col] = width;
            }
          }
        }
      }
    }

    /* parse custom row height
      example XML content
      <row r="1" spans="1:2" ht="44" customHeight="1" x14ac:dyDescent="0.15">
    */
    results = worksheet.findAllElements("row");
    if (results.isNotEmpty) {
      for (var element in results) {
        String? rowAttribute =
            element.getAttribute("r"); // i think min refers to the column
        String? heightAttribute = element.getAttribute("ht");
        if (rowAttribute != null && heightAttribute != null) {
          int? row = int.tryParse(rowAttribute);
          double? height = double.tryParse(heightAttribute);
          if (row != null && height != null) {
            row -= 1; // first col in _rowHeights is index 0
            if (row >= 0) {
              sheetObject._rowHeights[row] = height;
            }
          }
        }
      }
    }
  }
}

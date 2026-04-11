part of '../../excel_plus.dart';

/// Base class containing Parser fields, XML bootstrapping, and shared utilities.
///
/// Not meant to be used directly. Use [Parser] instead.
abstract class _ParserBase {
  final Excel _excel;
  final List<String> _rId = [];
  final Map<String, String> _worksheetTargets = {};

  _ParserBase(this._excel);

  void _parseContent({bool run = true});

  void _normalizeTable(Sheet sheet) {
    if (sheet._maxRows == 0 || sheet._maxColumns == 0) {
      sheet._sheetData.clear();
    }
    sheet._countRowsAndColumns();
  }

  void _putContentXml() {
    var file = _excel._archive.findFile("[Content_Types].xml");

    if (file == null) {
      _damagedExcel();
    }
    file!.decompress();
    _excel._xmlFiles["[Content_Types].xml"] =
        XmlDocument.parse(utf8.decode(file.content));
  }

  void _parseRelations() {
    var relations = _excel._archive.findFile('xl/_rels/workbook.xml.rels');
    if (relations != null) {
      relations.decompress();
      var document = XmlDocument.parse(utf8.decode(relations.content));
      _excel._xmlFiles['xl/_rels/workbook.xml.rels'] = document;

      document.findAllElements('Relationship').forEach((node) {
        String? id = node.getAttribute('Id');
        String? target = node.getAttribute('Target');
        if (target != null) {
          switch (node.getAttribute('Type')) {
            case _relationshipsStyles:
              _excel._stylesTarget = target;
              break;
            case _relationshipsWorksheet:
              if (id != null) _worksheetTargets[id] = target;
              break;
            case _relationshipsSharedStrings:
              _excel._sharedStringsTarget = target;
              break;
          }
        }
        if (id != null && !_rId.contains(id)) {
          _rId.add(id);
        }
      });
    } else {
      _damagedExcel();
    }
  }

  void _parseSharedStrings() {
    var sharedStrings =
        _excel._archive.findFile(_excel._absSharedStringsTarget);
    if (sharedStrings == null) {
      _excel._sharedStringsTarget = 'sharedStrings.xml';

      /// Running it with false will collect all the `rid` and will
      /// help us to get the available rid to assign it to `sharedStrings.xml` back
      _parseContent(run: false);

      if (_excel._xmlFiles.containsKey("xl/_rels/workbook.xml.rels")) {
        int rIdNumber = _getAvailableRid();

        _excel._xmlFiles["xl/_rels/workbook.xml.rels"]
            ?.findAllElements('Relationships')
            .first
            .children
            .add(XmlElement(
              XmlName('Relationship'),
              <XmlAttribute>[
                XmlAttribute(XmlName('Id'), 'rId$rIdNumber'),
                XmlAttribute(XmlName('Type'),
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'),
                XmlAttribute(XmlName('Target'), 'sharedStrings.xml')
              ],
            ));
        if (!_rId.contains('rId$rIdNumber')) {
          _rId.add('rId$rIdNumber');
        }
        String content =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        bool contain = true;

        _excel._xmlFiles["[Content_Types].xml"]
            ?.findAllElements('Override')
            .forEach((node) {
          var value = node.getAttribute('ContentType');
          if (value == content) {
            contain = false;
          }
        });
        if (contain) {
          _excel._xmlFiles["[Content_Types].xml"]
              ?.findAllElements('Types')
              .first
              .children
              .add(XmlElement(
                XmlName('Override'),
                <XmlAttribute>[
                  XmlAttribute(XmlName('PartName'), '/xl/sharedStrings.xml'),
                  XmlAttribute(XmlName('ContentType'), content),
                ],
              ));
        }
      }

      var content = utf8.encode(
          "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"/>");
      _excel._archive.addFile(
          ArchiveFile("xl/sharedStrings.xml", content.length, content));
      sharedStrings = _excel._archive.findFile("xl/sharedStrings.xml");
    }
    sharedStrings!.decompress();
    var document = XmlDocument.parse(utf8.decode(sharedStrings.content));
    _excel._xmlFiles["xl/${_excel._sharedStringsTarget}"] = document;

    document.findAllElements('si').forEach((node) {
      _parseSharedString(node);
    });
  }

  void _parseSharedString(XmlElement node) {
    final sharedString = SharedString(node: node);
    _excel._sharedStrings.add(sharedString, sharedString.stringValue);
  }

  dynamic _nodeChildren(XmlElement node, String child, {var attribute}) {
    Iterable<XmlElement> ele = node.findElements(child);
    if (ele.isNotEmpty) {
      if (attribute != null) {
        var attr = ele.first.getAttribute(attribute);
        if (attr != null) {
          return attr;
        }
        return null;
      }
      return true;
    }
    return null;
  }

  int _getFontIndex(XmlElement node, String text) {
    String? applyFont = node.getAttribute(text)?.trim();
    if (applyFont != null) {
      try {
        return int.parse(applyFont.toString());
      } catch (e) {
        if (applyFont.toLowerCase() == 'true') {
          return 1;
        }
      }
    }
    return 0;
  }

  int _getAvailableRid() {
    _rId.sort((a, b) {
      return int.parse(a.substring(3)).compareTo(int.parse(b.substring(3)));
    });

    List<String> got = List<String>.from(_rId.last.split(''));
    got.removeWhere((item) {
      return !'0123456789'.split('').contains(item);
    });
    return int.parse(got.join().toString()) + 1;
  }
}

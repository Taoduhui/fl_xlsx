part of 'xlsx_core.dart';

class SheetCell with XmlNodeWrapper {
  ExcelFile file;

  SheetCell(this.file, XmlNode node) {
    this.node = node;
  }

  WorkSheet get worksheet => file._docRefs[node.sheetDoc.path] as WorkSheet;

  String get reference => node.getAttribute("*r") ?? "";

  int get style => node.getAttribute("*s").asIntOrDefault(0);

  SheetCellType get type =>
      SheetCellTypeExt.parse(node.getAttribute("*t") ?? "n")!;

  int get cellMetadataIndex => node.getAttribute("*cm").asIntOrDefault(0);

  int get valueMetadataIndex => node.getAttribute("*vm").asIntOrDefault(0);

  bool get showPhonetic => node.getAttribute("*ph").asBoolOrFalse();

  String? get formula => node
      .into(
          selector: (c) =>
              c.type == XmlElementType.start && c.name.removeNamespace() == "f")
      ?.value;

  String? get value => node
      .into(
          selector: (c) =>
              c.type == XmlElementType.start && c.name.removeNamespace() == "v")
      ?.value;

  /// Rich Text String
  String? get inlineString => node
      .into(
          selector: (c) =>
              c.type == XmlElementType.start &&
              c.name.removeNamespace() == "is")
      ?.innerXML;

  bool get isMergedCell => mergeCell != null;

  MergeCell? get mergeCell {
    return worksheet.mergeCells?.firstWhere((c) => c.range().isInRange(CellRef(reference)));
  }

  String getString(){
    switch (type) {
      case SheetCellType.bool:
      // TODO: Handle this case.
        break;
      case SheetCellType.date:
      // TODO: Handle this case.
        break;
      case SheetCellType.error:
      // TODO: Handle this case.
        break;
      case SheetCellType.inlineString:
        return node.innerXML;
      case SheetCellType.number:
        return value!;
      case SheetCellType.sharedString:
        return _decodeSharedString();
      case SheetCellType.formulaString:
      // TODO: Handle this case.
        break;
    }
    return "";
  }

  String _decodeSharedString() {
    var index = value.asIntOrNull()!;
    var item = file.workbook._sharedStringTable.at(index);
    return item.text?.val ?? "";
  }
}

class Coordinate {
  static int minColumns = 1;
  static int maxColumns = 16384;

  static String columnNumberToName(int num) {
    if (num < minColumns || num > maxColumns) {
      throw Exception('Invalid column number');
    }
    String col = '';
    while (num > 0) {
      col = String.fromCharCode(((num - 1) % 26 + 65)) + col;
      num = (num - 1) ~/ 26;
    }
    return col;
  }

  static int columnNameToNumber(String name) {
    if (name.isEmpty) {
      throw Exception('Invalid column name');
    }
    int col = 0;
    int multi = 1;
    for (int i = name.length - 1; i >= 0; i--) {
      String char = name[i];
      if (char.codeUnitAt(0) >= 65 && char.codeUnitAt(0) <= 90) {
        col += (char.codeUnitAt(0) - 64) * multi;
      } else if (char.codeUnitAt(0) >= 97 && char.codeUnitAt(0) <= 122) {
        col += (char.codeUnitAt(0) - 96) * multi;
      } else {
        throw Exception('Invalid column name');
      }
      multi *= 26;
    }
    if (col > maxColumns) {
      throw Exception('Invalid column name');
    }
    return col;
  }
}

class CellRef {
  late String _raw;

  CellRef(this._raw);

  CellRef.fromIdx({int col = 1, int row = 1}) {
    _raw = "${Coordinate.columnNumberToName(col)}$row";
  }

  bool get isColFixed => _raw[0] == "\$";

  set isColFixed(bool v) {
    if (!v && isColFixed) {
      _raw = _raw.substring(1);
    } else if (v && !isColFixed) {
      _raw = "\$$_raw";
    }
  }

  String get col =>
      _raw.substring(0, _raw.indexOf(RegExp("[0-9]"))).replaceAll("\$", "");

  set col(String v) => _raw = _raw.replaceAll(col, v);

  bool get isRowFixed => _raw[_raw.indexOf(RegExp("[0-9]")) - 1] == "\$";

  set isRowFixed(bool v) {
    if (!v && isRowFixed) {
      var idx = _raw.indexOf(RegExp("[0-9]"));
      _raw = _raw.remove(idx - 1, idx);
    }
    if (v && !isRowFixed) {
      var idx = _raw.indexOf(RegExp("[0-9]"));
      _raw = _raw.insertBefore("\$", idx);
    }
  }

  String get row =>
      _raw.substring(_raw.indexOf(RegExp("[0-9]"))).replaceAll("\$", "");

  set row(String v) => _raw = _raw.replaceAll(row, v);

  int get colIdx => Coordinate.columnNameToNumber(col);

  set colIdx(int v) => col = Coordinate.columnNumberToName(v);

  int get rowIdx => int.parse(row);

  set rowIdx(int v) {
    if (v == 0) {
      throw Exception('Invalid row idx');
    }
    row = "$v";
  }

  @override
  String toString() {
    return _raw;
  }
}

class CellRefRange {
  String raw;

  CellRefRange([this.raw = "A1:A1"]);

  String get from => raw.contains(":") ? raw._beforeColon()! : raw;

  set from(String v) => raw = "$v:$to";

  String get to => raw.contains(":") ? raw._afterColon() : from;

  set to(String v) => raw = "$from:$v";

  bool isInRange(CellRef ref) {
    var fromRef = CellRef(from);
    var toRef = CellRef(to);
    if (ref.colIdx >= fromRef.colIdx &&
        ref.colIdx <= toRef.colIdx &&
        ref.rowIdx >= fromRef.rowIdx &&
        ref.rowIdx <= toRef.rowIdx) {
      return true;
    }
    return false;
  }
}

extension _CellRefStringExtension on String {
  String? _beforeColon() {
    var idx = indexOf(":");
    if (idx == -1) {
      return null;
    }
    return substring(0, idx);
  }

  String _afterColon() {
    var idx = indexOf(":");
    if (idx == -1) {
      return this;
    }
    return substring(idx + 1);
  }
}

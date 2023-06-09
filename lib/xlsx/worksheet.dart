part of 'xlsx_core.dart';

/// Base on ISO/IEC 29500-1:2016 Page 3912 PDF 3922
class WorkSheet {
  ExcelFile file;
  SheetDocument doc;
  late XmlNode node;

  final Map<String, SheetCell> _cells = {};

  final Map<int, SheetRow> _rows = {};

  WorkSheet(this.file, this.doc) {
    node = doc.root;
  }

  SheetProperties? get sheetProperties {
    var child = node.into(
        selector: (c) =>
            c.type == XmlElementType.start &&
            c.name.removeNamespace() == "sheetPr");
    if (child == null) {
      return null;
    }
    return SheetProperties(file, child);
  }

  SheetDimension? get dimension {
    var child = node.into(
        selector: (c) =>
            c.type == XmlElementType.start &&
            c.name.removeNamespace() == "dimension");
    if (child == null) {
      return null;
    }
    return SheetDimension(file, child);
  }

  SheetFormatProperties? get sheetFormatProperties {
    var child = node.into(
        selector: (c) =>
            c.type == XmlElementType.start &&
            c.name.removeNamespace() == "sheetFormatPr");
    if (child == null) {
      return null;
    }
    return SheetFormatProperties(file, child);
  }

  SheetColumns? get columns {
    var child = node.into(
        selector: (c) =>
            c.type == XmlElementType.start &&
            c.name.removeNamespace() == "cols");
    if (child == null) {
      return null;
    }
    return SheetColumns(file, child);
  }

  SheetData get sheetData {
    var child = node.into(
        selector: (c) =>
            c.type == XmlElementType.start &&
            c.name.removeNamespace() == "sheetData");
    if (child == null) {
      throw Exception("sheetData not found");
    }
    return SheetData(file, child);
  }

//TODO: CT_SheetCalcPr
//TODO: CT_SheetProtection
//TODO: CT_ProtectedRanges
//TODO: CT_Scenarios
//TODO: CT_AutoFilter
//TODO: CT_SortState
//TODO: CT_DataConsolidate
//TODO: CT_CustomSheetViews

  MergeCells? get mergeCells {
    var child = node.into(
        selector: (c) =>
            c.type == XmlElementType.start &&
            c.name.removeNamespace() == "mergeCells");
    if (child == null) {
      return null;
    }
    return MergeCells(file, child);
  }

//TODO: CT_PhoneticPr
//TODO: CT_ConditionalFormatting
//TODO: CT_DataValidations
//TODO: CT_Hyperlinks
//TODO: CT_PrintOptions
//TODO: CT_PageMargins
//TODO: CT_PageSetup
//TODO: CT_HeaderFooter
//TODO: CT_PageBreak
//TODO: CT_PageBreak
//TODO: CT_CustomProperties
//TODO: CT_CellWatches
//TODO: CT_IgnoredErrors
//TODO: CT_SmartTags
//TODO: CT_Drawing

  /////////////////////////////////////////

  CellRefRange get range {
    var cellRange = CellRefRange();
    int rMin = -1, rMax = -1;
    int cMin = -1, cMax = -1;
    for (var row in sheetData) {
      var rIdx = row.index;
      if (rMin == -1 || rMax == -1) {
        rMin = rIdx;
        rMax = rIdx;
      }
      if (rIdx < rMin) {
        rMin = rIdx;
      }
      if (rIdx > rMax) {
        rMax = rIdx;
      }
      for (var cell in row) {
        var ref = CellRef(cell.reference);
        var cIdx = ref.colIdx;
        if (cMin == -1 || cMax == -1) {
          cMin = cIdx;
          cMax = cIdx;
        }
        if (cIdx < cMin) {
          cMin = cIdx;
        }
        if (cIdx > cMax) {
          cMax = cIdx;
        }
      }
    }
    cellRange.from = CellRef.fromIdx(col: cMin, row: rMin).toString();
    cellRange.to = CellRef.fromIdx(col: cMax, row: rMax).toString();
    return cellRange;
  }

  SheetRow? rowAt(int idx) {
    if (_rows.isNotEmpty) {
      if (_rows.containsKey(idx)) {
        return _rows[idx];
      }
    }
    return sheetData.firstWhereOrNull((r) => r.index == idx);
  }

  SheetCell? cellAt(int col, int row, [SheetRow? rowNode]) {
    if (_cells.isNotEmpty) {
      var name = CellRef.fromIdx(col: col, row: row).toString();
      if (_cells.containsKey(name)) {
        return _cells[name];
      }
    }else if (rowNode != null) {
      return rowNode
          .firstWhereOrNull((cell) => CellRef(cell.reference).colIdx == col);
    } else if (_rows.isNotEmpty) {
      if (_rows.containsKey(row)) {
        return _rows[row]
            ?.firstWhereOrNull((cell) => CellRef(cell.reference).colIdx == col);
      }
    }
    return sheetData
        .firstWhereOrNull((r) => r.index == row)
        ?.firstWhereOrNull((cell) => CellRef(cell.reference).colIdx == col);
  }

  void enableCellCache() {
    for (var row in sheetData) {
      for (var cell in row) {
        _cells[cell.reference] = cell..mount();
      }
    }
  }

  void disableCellCache() {
    for (var cell in _cells.values) {
      cell.unmount();
    }
    _cells.clear();
  }

  void enableRowCache() {
    for (var row in sheetData) {
      _rows[row.index] = row..mount();
    }
  }

  void disableRowCache() {
    for (var row in _rows.values) {
      row.unmount();
    }
    _rows.clear();
  }
}

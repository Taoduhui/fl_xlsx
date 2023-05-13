part of 'xlsx_core.dart';

/// Base on ISO/IEC 29500-1:2016 Page 3912 PDF 3922
class WorkSheet {
  ExcelFile file;
  SheetDocument doc;
  late XmlNode node;

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

  SheetData? get sheetData {
    var child = node.into(
        selector: (c) =>
            c.type == XmlElementType.start &&
            c.name.removeNamespace() == "sheetData");
    if (child == null) {
      return null;
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

  SheetCell? cellAt(int col, int row) {
    return sheetData
        ?.firstWhereOrNull((r) => r.index == row)
        ?.firstWhereOrNull((cell) => CellRef(cell.reference).colIdx == col);
  }
}

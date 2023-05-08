part of 'xlsx_core.dart';

class SheetFormatProperties with XmlNodeWrapper {
  ExcelFile file;

  SheetFormatProperties(this.file,XmlNode node) {
    this.node = node;
  }

  int get baseColumnWidth => node.getAttribute("baseColWidth").asIntOrDefault(8);
  double? get defaultColumnWidth => node.getAttribute("defaultColWidth").asDoubleOrNull();
  double get defaultRowHeight => node.getAttribute("defaultRowHeight").asDoubleOrThrow();
  bool get customHeight => node.getAttribute("customHeight").asBoolOrDefault(false);
  bool get zeroHeight => node.getAttribute("zeroHeight").asBoolOrDefault(false);
  bool get thickTop => node.getAttribute("thickTop").asBoolOrDefault(false);
  bool get thickBottom => node.getAttribute("thickBottom").asBoolOrDefault(false);
  int get outlineLevelRow => node.getAttribute("outlineLevelRow").asIntOrDefault(0);
  int get outlineLevelColumn => node.getAttribute("outlineLevelCol").asIntOrDefault(0);
}
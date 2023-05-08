part of 'xlsx_core.dart';

class Selection with XmlNodeWrapper {
  ExcelFile file;
  Selection(this.file,XmlNode node) {
    this.node = node;
  }

  PaneType get pane => PaneTypeExt.parse(node.getAttribute("pane")??"topLeft");
  String? get activeCell => node.getAttribute("activeCell");
  int get activeCellId => node.getAttribute("activeCellId").asIntOrDefault(0);
  String get sqref => node.getAttribute("sqref")??"A1";
}
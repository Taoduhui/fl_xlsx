part of 'xlsx_core.dart';

class Pane with XmlNodeWrapper{
  ExcelFile file;
  Pane(this.file,XmlNode node){
    this.node = node;
  }

  double get xSplit => node.getAttribute("xSplit").asDoubleOrDefault(0.0);
  double get ySplit => node.getAttribute("ySplit").asDoubleOrDefault(0.0);
  String? get topLeftCell => node.getAttribute("topLeftCell");
  PaneType get activePane => PaneTypeExt.parse(node.getAttribute("activePane")??"topLeft");
  PaneState get state => PaneStateExt.parse(node.getAttribute("state")??"split");
}

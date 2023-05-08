part of 'xlsx_core.dart';

class SheetViewsIterator extends XmlChildNodeIterator<SheetView>{
  ExcelFile file;
  SheetViewsIterator(this.file,super.parent);

  @override
  SheetView build(n)=>SheetView(file,n);

  @override
  bool selector(n)=>n.type == XmlElementType.Normal && n.name.removeNamespace() == "sheetView";
}

class SheetViews with XmlNodeWrapper{
  SheetViews(XmlNode node){
    this.node = node;
  }
}

class SheetView with XmlNodeWrapper{
  ExcelFile file;

  SheetView(this.file,XmlNode node){
    this.node = node;
  }

  bool get windowProtection => node.getAttribute("windowProtection").asBoolOrFalse();
  bool get showFormulas => node.getAttribute("showFormulas").asBoolOrFalse();
  bool get showGridLines => node.getAttribute("showGridLines").asBoolOrTrue();
  bool get showRowColHeaders => node.getAttribute("showRowColHeaders").asBoolOrTrue();
  bool get showZeros => node.getAttribute("showZeros").asBoolOrTrue();
  bool get rightToLeft => node.getAttribute("rightToLeft").asBoolOrFalse();
  bool get tabSelected => node.getAttribute("tabSelected").asBoolOrFalse();
  bool get showRuler => node.getAttribute("showRuler").asBoolOrTrue();
  bool get showOutlineSymbols => node.getAttribute("showOutlineSymbols").asBoolOrTrue();
  bool get defaultGridColor => node.getAttribute("defaultGridColor").asBoolOrTrue();
  bool get showWhiteSpace => node.getAttribute("showWhiteSpace").asBoolOrTrue();
  SheetViewType get viewType =>  SheetViewTypeExt.parse(node.getAttribute("view")??"normal")!;
  String? get topLeftCell => node.getAttribute("topLeftCell");
  int get colorId => node.getAttribute("colorId").asIntOrDefault(64);
  int get zoomScale => node.getAttribute("zoomScale").asIntOrDefault(100);
  int get zoomScaleNormal => node.getAttribute("zoomScaleNormal").asIntOrDefault(0);
  int get zoomScaleSheetLayoutView => node.getAttribute("zoomScaleSheetLayoutView").asIntOrDefault(0);
  int get zoomScalePageLayoutView => node.getAttribute("zoomScalePageLayoutView").asIntOrDefault(0);
  int get workbookViewId => node.getAttribute("workbookViewId").asIntOrNull()!;

  Pane? get pane{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "pane");
    if(_node != null){
      return Pane(file,_node);
    }
    return null;
  }

  Selection? get selection{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "selection");
    if(_node != null){
      return Selection(file,_node);
    }
    return null;
  }

  PivotSelection? get pivotSelection{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "pivotSelection");
    if(_node != null){
      return PivotSelection(file,_node);
    }
    return null;
  }
}

part of 'xlsx_core.dart';

class PivotSelection with XmlNodeWrapper {
  ExcelFile file;

  PivotSelection(this.file,XmlNode node) {
    this.node = node;
  }

  PaneType get pane => PaneTypeExt.parse(node.getAttribute("*pane") ?? "topLeft");
  bool get showHeader => node.getAttribute("*showHeader").asBoolOrFalse();
  bool get label => node.getAttribute("*label").asBoolOrFalse();
  bool get data => node.getAttribute("*data").asBoolOrFalse();
  bool get extendable => node.getAttribute("*extendable").asBoolOrFalse();
  int get count => node.getAttribute("*count").asIntOrDefault(0);
  int get dimension => node.getAttribute("*dimension").asIntOrDefault(0);
  AxisType? get axis => AxisTypeExt.parse(node.getAttribute("*axis")??"");
  int get start => node.getAttribute("*start").asIntOrDefault(0);
  int get min => node.getAttribute("*min").asIntOrDefault(0);
  int get max => node.getAttribute("*max").asIntOrDefault(0);
  int get activeRow => node.getAttribute("*activeRow").asIntOrDefault(0);
  int get activeCol => node.getAttribute("*activeCol").asIntOrDefault(0);
  int get previousRow => node.getAttribute("*previousRow").asIntOrDefault(0);
  int get previousCol => node.getAttribute("*previousCol").asIntOrDefault(0);
  int get click => node.getAttribute("*click").asIntOrDefault(0);
  String? get refId => node.getAttribute("*r:id");

  PivotArea? get references{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "pivotArea");
    if(cnode != null){
      return PivotArea(file,cnode);
    }
    return null;
  }
}



class PivotArea with XmlNodeWrapper {
  ExcelFile file;

  PivotArea(this.file,XmlNode node) {
    this.node = node;
  }

  int? get field => node.getAttribute("*field").asIntOrNull();
  PivotAreaType get type => PivotAreaTypeExt.parse(node.getAttribute("*type") ?? "normal")!;
  bool get dataOnly => node.getAttribute("*dataOnly").asBoolOrTrue();
  bool get labelOnly => node.getAttribute("*labelOnly").asBoolOrFalse();
  bool get grandRow => node.getAttribute("*grandRow").asBoolOrFalse();
  bool get grandCol => node.getAttribute("*grandCol").asBoolOrFalse();
  bool get cacheIndex => node.getAttribute("*cacheIndex").asBoolOrFalse();
  bool get outline => node.getAttribute("*outline").asBoolOrTrue();
  String? get offset => node.getAttribute("*offset");
  bool get collapsedLevelsAreSubtotals => node.getAttribute("*collapsedLevelsAreSubtotals").asBoolOrFalse();
  AxisType? get axis => AxisTypeExt.parse(node.getAttribute("*axis")??"");
  int? get fieldPosition => node.getAttribute("*fieldPosition").asIntOrNull();

  PivotAreaReferences? get references{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "references");
    if(cnode != null){
      return PivotAreaReferences(file,cnode);
    }
    return null;
  }
}

class PivotAreaReferencesIterator extends XmlChildNodeIterator<PivotAreaReference>{
  ExcelFile file;

  PivotAreaReferencesIterator(this.file,super.parent);

  @override
  PivotAreaReference build(n)=>PivotAreaReference(file,n);

  @override
  bool selector(n)=>n.type == XmlElementType.start && n.name.removeNamespace() == "reference";
}

class PivotAreaReferences extends Iterable<PivotAreaReference> with XmlNodeWrapper {
  ExcelFile file;

  PivotAreaReferences(this.file,XmlNode node) {
    this.node = node;
  }

  @override
  Iterator<PivotAreaReference> get iterator => PivotAreaReferencesIterator(file,node);

  int get count => node.getAttribute("*count").asIntOrNull()!;
}

class PivotAreaReferenceIterator extends XmlChildNodeIterator<IndexType>{
  ExcelFile file;

  PivotAreaReferenceIterator(this.file,super.parent);

  @override
  IndexType build(n)=>IndexType(file,n);

  @override
  bool selector(n)=>n.type == XmlElementType.start && n.name.removeNamespace() == "x";
}

class PivotAreaReference extends Iterable<IndexType> with XmlNodeWrapper {
  ExcelFile file;

  PivotAreaReference(this.file,XmlNode node) {
    this.node = node;
  }

  int? get field => node.getAttribute("*field").asIntOrNull();
  int get count => node.getAttribute("*count").asIntOrNull()!;
  bool get selected => node.getAttribute("*selected").asBoolOrTrue();
  bool get byPosition => node.getAttribute("*byPosition").asBoolOrFalse();
  bool get relative => node.getAttribute("*relative").asBoolOrFalse();
  bool get defaultSubtotal => node.getAttribute("*defaultSubtotal").asBoolOrFalse();
  bool get sumSubtotal => node.getAttribute("*sumSubtotal").asBoolOrFalse();
  bool get countASubtotal => node.getAttribute("*countASubtotal").asBoolOrFalse();
  bool get avgSubtotal => node.getAttribute("*avgSubtotal").asBoolOrFalse();
  bool get maxSubtotal => node.getAttribute("*maxSubtotal").asBoolOrFalse();
  bool get minSubtotal => node.getAttribute("*minSubtotal").asBoolOrFalse();
  bool get productSubtotal => node.getAttribute("*productSubtotal").asBoolOrFalse();
  bool get countSubtotal => node.getAttribute("*countSubtotal").asBoolOrFalse();
  bool get stdDevSubtotal => node.getAttribute("*stdDevSubtotal").asBoolOrFalse();
  bool get stdDevPSubtotal => node.getAttribute("*stdDevPSubtotal").asBoolOrFalse();
  bool get varSubtotal => node.getAttribute("*varSubtotal").asBoolOrFalse();
  bool get varPSubtotal => node.getAttribute("*varPSubtotal").asBoolOrFalse();

  @override
  Iterator<IndexType> get iterator => PivotAreaReferenceIterator(file,node);
}

class IndexType with XmlNodeWrapper {
  ExcelFile file;

  IndexType(this.file,XmlNode node) {
    this.node = node;
  }

  int get value => node.getAttribute("*v").asIntOrNull()!;
}
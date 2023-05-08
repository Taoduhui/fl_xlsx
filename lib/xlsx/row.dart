
part of "./xlsx_core.dart";

class SheetRowIterator extends XmlChildNodeIterator<SheetCell>{
  ExcelFile file;
  SheetRowIterator(this.file,super.parent);

  @override
  SheetCell build(n)=>SheetCell(file,n);

  @override
  bool selector(n)=>n.type == XmlElementType.Normal && n.name.removeNamespace() == "c";
}

class SheetRow extends Iterable<SheetCell> with XmlNodeWrapper{

  ExcelFile file;

  SheetRow(this.file,XmlNode node){
    this.node = node;
  }

  int get index =>  int.parse(node.getAttribute("r")??"-1");

  String? get spans =>  node.getAttribute("spans");

  String? get height => node.getAttribute("ht");

  bool get customHeight => node.getAttribute("customHeight").asBoolOrFalse();

  bool get collapsed => node.getAttribute("collapsed").asBoolOrFalse();

  bool get hidden => node.getAttribute("hidden").asBoolOrFalse();

  int get outlineLevel =>  node.getAttribute("outlineLevel").asIntOrDefault(0);

  bool get showPhonetic => node.getAttribute("ph").asBoolOrFalse();

  int get styleIndex =>  node.getAttribute("s").asIntOrDefault(0);

  bool get customFormat => node.getAttribute("customFormat").asBoolOrFalse();

  bool get thickBottomBorder => node.getAttribute("thickBot").asBoolOrFalse();

  bool get thickTopBorder => node.getAttribute("thickTop").asBoolOrFalse();

  @override
  Iterator<SheetCell> get iterator => SheetRowIterator(file,node);
}
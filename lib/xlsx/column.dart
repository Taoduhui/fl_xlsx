part of 'xlsx_core.dart';

class SheetColumnsIterator extends XmlChildNodeIterator<SheetColumn>{
  ExcelFile file;
  SheetColumnsIterator(this.file,super.parent);

  @override
  SheetColumn build(XmlNode n) => SheetColumn(file,n);

  @override
  bool selector(XmlNode n) => n.type == XmlElementType.start && n.name.removeNamespace() == "col";
}

class SheetColumns extends Iterable<SheetColumn> with XmlNodeWrapper {
  ExcelFile file;

  SheetColumns(this.file,XmlNode node) {
    this.node = node;
  }

  @override
  Iterator<SheetColumn> get iterator => SheetColumnsIterator(file,node);
}

class SheetColumn with XmlNodeWrapper {
  ExcelFile file;

  SheetColumn(this.file,XmlNode node) {
    this.node = node;
  }

  int get min => node.getAttribute("*min").asIntOrThrow();
  int get max => node.getAttribute("*max").asIntOrThrow();
  double? get width => node.getAttribute("*width").asDoubleOrNull();
  int get style => node.getAttribute("*style").asIntOrDefault(0);
  bool get hidden => node.getAttribute("*hidden").asBoolOrDefault(false);
  bool get bestFit => node.getAttribute("*bestFit").asBoolOrDefault(false);
  bool get customWidth => node.getAttribute("*customWidth").asBoolOrDefault(false);
  bool get phonetic => node.getAttribute("*phonetic").asBoolOrDefault(false);
  int get outlineLevel => node.getAttribute("*outlineLevel").asIntOrDefault(0);
  bool get collapsed => node.getAttribute("*collapsed").asBoolOrDefault(false);
}

part of 'xlsx_core.dart';

class MergeCellsIterator extends XmlChildNodeIterator<MergeCell>{
  ExcelFile file;
  MergeCellsIterator(this.file,super.parent);

  @override
  MergeCell build(XmlNode n) => MergeCell(file,n);

  @override
  bool selector(XmlNode n) => n.type == XmlElementType.start && n.name.removeNamespace() == "col";
}

class MergeCells extends Iterable<MergeCell> with XmlNodeWrapper {
  ExcelFile file;

  MergeCells(this.file,XmlNode node) {
    this.node = node;
  }

  @override
  Iterator<MergeCell> get iterator => MergeCellsIterator(file,node);
}

class MergeCell with XmlNodeWrapper{
  ExcelFile file;

  MergeCell(this.file,XmlNode node){
    this.node = node;
  }

  String get ref => node.getAttribute("*ref")!;

  String get from => ref._beforeColon()!;
  set from(String v) => node.getAttributeNode("*ref")!.setAttribute("$v:$to");

  String get to => ref._afterColon();
  set to(String v) => node.getAttributeNode("*ref")!.setAttribute("$from:$v");

  CellRefRange range() => CellRefRange(ref);
}
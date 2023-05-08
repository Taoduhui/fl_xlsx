part of 'xlsx_core.dart';

class SheetDataIterator extends XmlChildNodeIterator<SheetRow>{
  ExcelFile file;
  SheetDataIterator(this.file,super.parent);

  @override
  SheetRow build(XmlNode n) => SheetRow(file,n);

  @override
  bool selector(XmlNode n) => n.type == XmlElementType.Normal && n.name.removeNamespace() == "row";
}

class SheetData extends Iterable<SheetRow> with XmlNodeWrapper {
  ExcelFile file;

  SheetData(this.file,XmlNode node) {
    this.node = node;
  }

  @override
  Iterator<SheetRow> get iterator => SheetDataIterator(file,node);
}
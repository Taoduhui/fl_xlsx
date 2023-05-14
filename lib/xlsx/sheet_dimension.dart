part of 'xlsx_core.dart';

class SheetDimension with XmlNodeWrapper{
  ExcelFile file;

  SheetDimension(this.file,XmlNode node){
    this.node = node;
  }

  String get reference => node.getAttribute("*ref") ?? "";
}

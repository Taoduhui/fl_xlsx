part of 'xlsx_core.dart';

class OutlineProperty with XmlNodeWrapper{
  ExcelFile file;

  OutlineProperty(this.file,XmlNode node){
    this.node = node;
  }

  bool get applyStyles => node.getAttribute("applyStyles").asBoolOrFalse();

  bool get summaryBelow => node.getAttribute("summaryBelow").asBoolOrTrue();

  bool get summaryRight => node.getAttribute("summaryRight").asBoolOrTrue();

  bool get showOutlineSymbols => node.getAttribute("showOutlineSymbols").asBoolOrTrue();
}
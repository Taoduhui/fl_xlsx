part of './xlsx_core.dart';

class PageSetupProperty with XmlNodeWrapper{
  ExcelFile file;

  PageSetupProperty(this.file,XmlNode node){
    this.node = node;
  }

  bool get autoPageBreaks => node.getAttribute("autoPageBreaks").asBoolOrTrue();

  bool get fitToPage => node.getAttribute("fitToPage").asBoolOrFalse();
}
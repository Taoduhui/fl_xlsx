part of 'xlsx_core.dart';

class SheetsIterator extends XmlChildNodeIterator<Sheet>{
  ExcelFile file;

  SheetsIterator(this.file,super.parent);

  @override
  Sheet build(XmlNode n) => Sheet(file,n);

  @override
  bool selector(XmlNode n) => n.type == XmlElementType.Normal && n.name.removeNamespace() == "sheet";

}

class Sheets extends Iterable<Sheet> with XmlNodeWrapper{
  ExcelFile file;

  Sheets(this.file,XmlNode node){
    this.node = node;
  }

  @override
  Iterator<Sheet> get iterator => SheetsIterator(file,node);
}

enum SheetState {
  visible,
  hidden,
  veryHidden
}

extension SheetStateExt on SheetState {
  static SheetState? parse(String str) {
    switch (str) {
      case "visible":
        return SheetState.visible;
      case "hidden":
        return SheetState.hidden;
      case "veryHidden":
        return SheetState.veryHidden;
    }
  }

  String get key {
    switch (this) {
      case SheetState.visible:
        return "visible";
      case SheetState.hidden:
        return "hidden";
      case SheetState.veryHidden:
        return "veryHidden";
    }
  }
}

class Sheet with XmlNodeWrapper {
  ExcelFile file;

  Sheet(this.file,XmlNode node) {
    this.node = node;
  }

  String get name => node.getAttribute("name") ?? "";
  int get sheetId => node.getAttribute("sheetId").asIntOrNull()!;
  SheetState get state => SheetStateExt.parse(node.getAttribute("state") ?? "visible")!;
  String get rid => node.getAttribute("r:id") ?? "";

  WorkSheet get workSheet{
     var workSheetRel = file.workbook.relationships.firstWhere((rel) => rel.id == rid);
     var workSheetPath = workSheetRel.target.absolutePath(node.sheetDoc.path.directory);
     SheetDocument workSheetDoc = file.openFile(workSheetPath);
     return WorkSheet(file, workSheetDoc.root);
  }
}


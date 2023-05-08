part of 'xlsx_core.dart';

class SheetCell with XmlNodeWrapper{
  ExcelFile file;

  SheetCell(this.file,XmlNode node){
    this.node = node;
  }

  String get reference => node.getAttribute("r") ?? "";

  int get style => node.getAttribute("s").asIntOrDefault(0);

  SheetCellType get type => SheetCellTypeExt.parse(node.getAttribute("t") ?? "n")!;

  int get cellMetadataIndex => node.getAttribute("cm").asIntOrDefault(0);

  int get valueMetadataIndex => node.getAttribute("vm").asIntOrDefault(0);

  bool get showPhonetic => node.getAttribute("ph").asBoolOrFalse();

  String? get formula => node.into(selector: (c)=> c.type == XmlElementType.Normal && c.name.removeNamespace() == "f")?.value;

  String? get value => node.into(selector: (c)=> c.type == XmlElementType.Normal && c.name.removeNamespace() == "v")?.value;

  /// Rich Text String
  String? get inlineString => node.into(selector: (c)=> c.type == XmlElementType.Normal && c.name.removeNamespace() == "is")?.innerXML;

  String decode(){
    switch(type){
      case SheetCellType.bool:
        // TODO: Handle this case.
        break;
      case SheetCellType.date:
        // TODO: Handle this case.
        break;
      case SheetCellType.error:
        // TODO: Handle this case.
        break;
      case SheetCellType.inlineString:
        // TODO: Handle this case.
        break;
      case SheetCellType.number: return value!;
      case SheetCellType.sharedString: return _decodeSharedString();
      case SheetCellType.formulaString:
        // TODO: Handle this case.
        break;
    }
    return "";
  }

  String _decodeSharedString(){
    var index = value.asIntOrNull()!;
    var item = file.workbook.sharedStringTable.at(index);
    return item.text?.val ?? "";
  }
}
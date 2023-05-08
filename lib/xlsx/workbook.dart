part of "./xlsx_core.dart";

/// Base on ISO/IEC 29500-1:2016 Page 3948 PDF 3958
class WorkBook with XmlNodeWrapper{

  late Relationships relationships;
  late SharedStringTable sharedStringTable;
  ExcelFile file;

  WorkBook(this.file,XmlNode node){
    this.node = node;
    var workbookDir = node.sheetDoc.path.directory;
    var bookRelsPath = "${workbookDir}/_rels/workbook.xml.rels";
    SheetDocument bookRels = file.openFile(bookRelsPath);
    relationships = Relationships(file,bookRels.root);
    var sharedStringsRel = relationships.firstWhereOrNull((rel) => rel.type == RelationshipType.sharedStrings);
    if(sharedStringsRel != null){
      var sharedStringsPath = sharedStringsRel.target.absolutePath(workbookDir);
      SheetDocument sharedStringTableDoc = file.openFile(sharedStringsPath);
      sharedStringTable = SharedStringTable(file,sharedStringTableDoc.root);
    }
  }

  Sheets get sheets{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "sheets");
    if(_node != null){
      return Sheets(file,_node);
    }
    throw Exception("broken file: workbook no sheets");
  }

  //TODO: CT_FileVersion
  //TODO: CT_FileSharing
  //TODO: CT_WorkbookPr
  //TODO: CT_WorkbookProtection
  //TODO: CT_BookViews
  //TODO: CT_FunctionGroups
  //TODO: CT_ExternalReferences
  //TODO: CT_DefinedNames
  //TODO: CT_CalcPr
  //TODO: CT_OleSize
  //TODO: CT_CustomWorkbookViews
  //TODO: CT_PivotCaches
  //TODO: CT_SmartTagPr
  //TODO: CT_SmartTagTypes
  //TODO: CT_WebPublishing
  //TODO: CT_FileRecoveryPr
  //TODO: CT_WebPublishObjects
  //TODO: CT_ExtensionList
}
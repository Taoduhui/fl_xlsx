part of "./xlsx_core.dart";

/// Base on ISO/IEC 29500-1:2016 Page 3948 PDF 3958
class WorkBook with XmlNodeWrapper{

  late Relationships _relationships;
  late SharedStringTable _sharedStringTable;
  ExcelFile file;

  WorkBook(this.file,XmlNode node){
    this.node = node;
    var workbookDir = node.sheetDoc.path.directory;
    var bookRelsPath = "${workbookDir}/_rels/workbook.xml.rels";
    _relationships = file._openFile(bookRelsPath,(doc)=>Relationships(file,doc.root));
    var sharedStringsRel = _relationships.firstWhereOrNull((rel) => rel.type == RelationshipType.sharedStrings);
    if(sharedStringsRel != null){
      var sharedStringsPath = sharedStringsRel.target.absolutePath(workbookDir);
      _sharedStringTable = file._openFile(sharedStringsPath,(doc)=>SharedStringTable(file,doc.root));
    }
  }

  Sheets get sheets{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "sheets");
    if(cnode != null){
      return Sheets(file,cnode);
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
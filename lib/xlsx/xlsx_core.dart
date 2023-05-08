import 'dart:convert';
import 'dart:typed_data';

import 'package:archive/archive.dart';
import '../xml/xml_core.dart';

part 'relationship.dart';
part 'workbook.dart';
part 'shared_strings.dart';
part 'sheet_string_ext.dart';
part 'row.dart';
part 'cell.dart';
part 'sheet.dart';
part 'sheet_properties.dart';
part 'outline_property.dart';
part 'page_setup_property.dart';
part 'sheet_dimension.dart';
part 'sheet_views.dart';
part 'pane.dart';
part 'selection.dart';
part 'pivot_selection.dart';
part 'common_type.dart';
part 'sheet_format_properties.dart';
part 'column.dart';
part 'xml_utils.dart';
part 'sheet_data.dart';
part 'worksheet.dart';
part 'config.dart';

class ExcelFile{

  late Archive archive;
  late ExcelConfig config;

  ExcelFile.fromBytes(Uint8List bytes,{ExcelConfig? config}){
    archive = ZipDecoder().decodeBytes(bytes);
    this.config = config??ExcelConfig();
    _init();
  }

  Map<String,SheetDocument> documents = {};

  late Relationships relationships;
  late WorkBook workbook;

  void _init(){
    _initRelationships();
  }

  SheetDocument openFile(String path){
    if(!documents.containsKey(path)){
      var rel = archive.findFile(path);
      if(rel != null){
        rel.decompress();
        var doc = SheetDocument.fromString(utf8.decode(rel.content),path);
        documents[path] = doc;
      }else{
        throw Exception("file not found");
      }
    }
    return documents[path]!;
  }

  void _initRelationships(){
    var pkgRels = "_rels/.rels";
    var doc = openFile(pkgRels);
    relationships = Relationships(this,doc.root);
    var bookRels = relationships.where((r) => r.type == RelationshipType.officeDocument);
    if(bookRels.isEmpty) throw Exception("broken file");
    var bookRel = bookRels.first;
    _initWorkBook(bookRel.target.absolutePath(""));
  }

  void _initWorkBook(String path){
    var doc = openFile(path);
    workbook = WorkBook(this,doc.root);
  }
}
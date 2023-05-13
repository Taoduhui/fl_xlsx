import 'dart:convert';
import 'dart:typed_data';

import 'package:archive/archive.dart';
import 'package:large_xml/large_xml.dart';
// import '../xml/xml_core.dart';

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
part 'merge_cell.dart';

class ExcelFile{

  late Archive _archive;

  late ExcelConfig _config;
  ExcelConfig get config => _config;

  ExcelFile.fromBytes(Uint8List bytes,{ExcelConfig? config}){
    _archive = ZipDecoder().decodeBytes(bytes);
    _config = config??ExcelConfig();
    _init();
  }

  final Map<String,SheetDocument> _documents = {};
  final Map<String,dynamic> _docRefs = {};

  late Relationships _relationships;
  late WorkBook workbook;

  void _init(){
    _initRelationships();
  }

  T _openFile<T>(String path,T Function(SheetDocument doc) builder){
    if(!_documents.containsKey(path)){
      var rel = _archive.findFile(path);
      if(rel != null){
        rel.decompress();
        var doc = SheetDocument.fromString(utf8.decode(rel.content),path);
        _documents[path] = doc;
      }else{
        throw Exception("file not found");
      }
    }
    _docRefs[path] = builder(_documents[path]!);
    return _docRefs[path]!;
  }

  void _initRelationships(){
    var pkgRels = "_rels/.rels";
    _relationships = _openFile(pkgRels,(doc)=> Relationships(this,doc.root));
    var bookRels = _relationships.where((r) => r.type == RelationshipType.officeDocument);
    if(bookRels.isEmpty) throw Exception("broken file");
    var bookRel = bookRels.first;
    _initWorkBook(bookRel.target.absolutePath(""));
  }

  void _initWorkBook(String path){
    workbook = _openFile(path,(doc)=>WorkBook(this,doc.root));
  }
}
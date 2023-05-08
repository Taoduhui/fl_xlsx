
part of "./xlsx_core.dart";

class RelationshipsIterator extends XmlChildNodeIterator<Relationship>{
  RelationshipsIterator(super.parent);

  @override
  Relationship build(XmlNode n) => Relationship(n);

  @override
  bool selector(XmlNode n) => n.type == XmlElementType.Normal && n.name.removeNamespace() == "Relationship";
}

class Relationships extends Iterable<Relationship> with XmlNodeWrapper{

  ExcelFile file;

  Relationships(this.file,XmlNode node){
    this.node = node;
  }

  @override
  Iterator<Relationship> get iterator => RelationshipsIterator(node);

}

enum RelationshipType{
  vmlDrawing,
  image,
  table,
  comments,
  hyperlink,
  drawing,
  officeDocument,
  core,
  app,
  theme,
  styles,
  sharedStrings,
  worksheet,
}

extension RelationshipTypeExt on RelationshipType{

  static const vmlDrawing = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
  static const image = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
  static const table = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
  static const comments = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
  static const hyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
  static const drawing = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
  static const officeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
  static const core = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
  static const app = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
  static const theme =  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
  static const styles =  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
  static const sharedStrings =  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
  static const worksheet =  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

  static RelationshipType? parse(String str){
    switch(str){
      case vmlDrawing : return RelationshipType.vmlDrawing;
      case image : return RelationshipType.image;
      case table : return RelationshipType.table;
      case comments : return RelationshipType.comments;
      case hyperlink : return RelationshipType.hyperlink;
      case drawing : return RelationshipType.drawing;
      case officeDocument : return RelationshipType.officeDocument;
      case core : return RelationshipType.core;
      case app : return RelationshipType.app;
      case theme : return RelationshipType.theme;
      case styles : return RelationshipType.styles;
      case sharedStrings : return RelationshipType.sharedStrings;
      case worksheet : return RelationshipType.worksheet;
    }
    return null;
  }

  String get key{
    switch(this){
      case  RelationshipType.vmlDrawing : return vmlDrawing;
      case  RelationshipType.image : return image;
      case  RelationshipType.table : return table;
      case  RelationshipType.comments : return comments;
      case  RelationshipType.hyperlink : return hyperlink;
      case  RelationshipType.drawing : return drawing;
      case  RelationshipType.officeDocument : return officeDocument;
      case  RelationshipType.core : return core;
      case  RelationshipType.app : return app;
      case  RelationshipType.theme : return theme;
      case  RelationshipType.styles : return styles;
      case  RelationshipType.sharedStrings : return sharedStrings;
      case  RelationshipType.worksheet : return worksheet;
    }
  }
}

enum RelationshipTargetMode{
  external,
  internal
}

extension RelationshipTargetModeExt on RelationshipTargetMode{
  static RelationshipTargetMode? parse(String str){
    switch(str){
      case "External" : return RelationshipTargetMode.external;
      case "Internal" : return RelationshipTargetMode.internal;
    }
    return null;
  }

  String get key{
    switch(this){
      case  RelationshipTargetMode.external : return "External";
      case  RelationshipTargetMode.internal : return "Internal";
    }
  }
}

class Relationship with XmlNodeWrapper{

  Relationship(XmlNode node){
    this.node = node;
  }

  String get id => node.getAttribute("Id")??"";

  String get target => node.getAttribute("Target")??"";

  RelationshipType get type => RelationshipTypeExt.parse(node.getAttribute("Type")??"")!;

  RelationshipTargetMode get targetMode => RelationshipTargetModeExt.parse(node.getAttribute("TargetMode")??"")!;
}
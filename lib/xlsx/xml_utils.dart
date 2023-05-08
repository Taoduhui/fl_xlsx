part of 'xlsx_core.dart';

class SheetDocument extends XmlDocument{

  SheetDocument.fromString(super.raw,this.path) : super.fromString();

  String path;
}

extension XmlExtension on XmlNode{
  SheetDocument get sheetDoc => document as SheetDocument;
}

abstract class XmlChildNodeIterator<T> extends Iterator<T>{
  XmlNode parent;
  XmlNode? node;

  XmlChildNodeIterator(this.parent);

  T build(XmlNode n);

  bool selector(XmlNode n);

  @override
  T get current => build(node!);

  @override
  bool moveNext() {
    if(node == null){
      node = parent.into(selector: selector);
    }else{
      node = node!.next(selector: selector);
    }
    if(node != null){
      return true;
    }
    return false;
  }
}

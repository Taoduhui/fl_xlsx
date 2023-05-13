part of 'xlsx_core.dart';

extension _SheetStringExtension on String{

  String insertAfter(String str, int index) {
    var str1 = substring(0, index + 1);
    var str2 = substring(index + 1);
    return "$str1$str$str2";
  }

  String insertBefore(String str, int index) {
    var str1 = substring(0, index);
    var str2 = substring(index);
    return "$str1$str$str2";
  }

  String remove(int start, int end) {
    var str1 = substring(0, start);
    var str2 = substring(end);
    return "$str1$str2";
  }
}

extension _SheetIterableExtension<T> on Iterable<T>{
  T? firstWhereOrNull(bool Function(T element) test){
    for(var i in this){
      if(test(i)){
        return i;
      }
    }
    return null;
  }
}

mixin XmlNodeWrapper{
  late XmlNode node;

  void mount()=>node.mount();
  void unmount()=>node.unmount();
  bool get mounted=>node.mounted;
}

class XmlNodeIterator extends Iterator<XmlNode>{

  XmlNode node;
  XmlElementType? type;
  bool Function(XmlNode node)? selector;

  XmlNodeIterator(this.node,{this.type,this.selector});

  @override
  XmlNode get current => node;

  @override
  bool moveNext() {
    var next = node.next(type: type,selector: selector);
    if(next != null){
      node = next;
      return true;
    }
    return false;
  }
}

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

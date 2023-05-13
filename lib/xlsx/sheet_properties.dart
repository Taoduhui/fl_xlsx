part of 'xlsx_core.dart';

class SheetProperties with XmlNodeWrapper{
  ExcelFile file;

  SheetProperties(this.file,XmlNode node){
    this.node = node;
  }

  String? get codeName => node.into(selector: (c)=> c.type == XmlElementType.start && c.name.removeNamespace() == "codeName")?.value;

  bool get enableFormatConditionsCalculation => node.getAttribute("*enableFormatConditionsCalculation").asBoolOrTrue();

  bool get filterMode => node.getAttribute("*filterMode").asBoolOrFalse();

  bool get published => node.getAttribute("*published").asBoolOrTrue();

  bool get syncHorizontal => node.getAttribute("*syncHorizontal").asBoolOrFalse();

  bool get syncVertical => node.getAttribute("*syncVertical").asBoolOrFalse();

  String? get syncRef => node.into(selector: (c)=> c.type == XmlElementType.start && c.name.removeNamespace() == "syncRef")?.value;

  bool get enableTransitionEntry => node.getAttribute("*transitionEntry").asBoolOrFalse();

  bool get enableTransitionEvaluation => node.getAttribute("*transitionEvaluation").asBoolOrFalse();

  ColorProperty? get tabColor{
    var child =  node.into(selector: (c)=> c.type == XmlElementType.start && c.name.removeNamespace() == "tabColor");
    if(child == null){
      return null;
    }
    return ColorProperty(file,child);
  }

  OutlineProperty? get outlineProperty{
    var child =  node.into(selector: (c)=> c.type == XmlElementType.start && c.name.removeNamespace() == "outlinePr");
    if(child == null){
      return null;
    }
    return OutlineProperty(file,child);
  }

  PageSetupProperty? get pageSetupProperty{
    var child =  node.into(selector: (c)=> c.type == XmlElementType.start && c.name.removeNamespace() == "pageSetUpPr");
    if(child == null){
      return null;
    }
    return PageSetupProperty(file,child);
  }
}
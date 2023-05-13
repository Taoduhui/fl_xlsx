part of 'xlsx_core.dart';

enum SheetCellType{
  bool,//b
  date,//d
  error,//e
  inlineString,//inlineStr
  number,//n
  sharedString,//s
  formulaString,//str
}

extension SheetCellTypeExt on SheetCellType{
  String get key{
    switch(this){
      case SheetCellType.bool : return "b";
      case SheetCellType.date : return "d";
      case SheetCellType.error : return "e";
      case SheetCellType.inlineString : return "inlineStr";
      case SheetCellType.number : return "n";
      case SheetCellType.sharedString : return "s";
      case SheetCellType.formulaString : return "str";
    }
  }

  static SheetCellType? parse(String key){
    switch(key){
      case "b": return SheetCellType.bool;
      case "d": return SheetCellType.date;
      case "e": return SheetCellType.error;
      case "inlineStr": return SheetCellType.inlineString;
      case "n": return SheetCellType.number;
      case "s": return SheetCellType.sharedString;
      case "str": return SheetCellType.formulaString;
    }
    return null;
  }
}

enum SheetViewType{
  normal,
  pageBreakPreview,
  pageLayout
}

extension SheetViewTypeExt on SheetViewType{
  static SheetViewType? parse(String str){
    switch(str){
      case "normal" : return SheetViewType.normal;
      case "pageBreakPreview" : return SheetViewType.pageBreakPreview;
      case "pageLayout" : return SheetViewType.pageLayout;
    }
    return null;
  }

  String get key{
    switch(this){
      case SheetViewType.normal : return "normal";
      case SheetViewType.pageBreakPreview : return "pageBreakPreview";
      case SheetViewType.pageLayout : return "pageLayout";
    }
  }
}

enum AxisType { axisRow, axisCol, axisPage, axisValues }

extension AxisTypeExt on AxisType {
  static AxisType? parse(String str){
    switch(str){
      case "axisRow": return AxisType.axisRow;
      case "axisCol": return AxisType.axisCol;
      case "axisPage": return AxisType.axisPage;
      case "axisValues": return AxisType.axisValues;
    }
    return null;
  }

  String get value {
    switch(this){
      case AxisType.axisRow: return "axisRow";
      case AxisType.axisCol: return "axisCol";
      case AxisType.axisPage: return "axisPage";
      case AxisType.axisValues: return "axisValues";
    }
  }
}

enum PaneType {
  bottomRight,
  topRight,
  bottomLeft,
  topLeft,
}

extension PaneTypeExt on PaneType {
  static PaneType parse(String str){
    switch(str){
      case "bottomRight": return PaneType.bottomRight;
      case "topRight": return PaneType.topRight;
      case "bottomLeft": return PaneType.bottomLeft;
      case "topLeft": return PaneType.topLeft;
    }
    throw Exception("unknown pane type: $str");
  }

  String get key{
    switch(this){
      case PaneType.bottomRight: return "bottomRight";
      case PaneType.topRight: return "topRight";
      case PaneType.bottomLeft: return "bottomLeft";
      case PaneType.topLeft: return "topLeft";
    }
  }
}

enum PaneState{
  split,
  frozen,
  frozenSplit,
}

extension PaneStateExt on PaneState{
  static PaneState parse(String str){
    switch(str){
      case "split": return PaneState.split;
      case "frozen": return PaneState.frozen;
      case "frozenSplit": return PaneState.frozenSplit;
    }
    throw Exception("unknown pane state:$str");
  }

  String get key{
    switch(this){
      case PaneState.split: return "split";
      case PaneState.frozen: return "frozen";
      case PaneState.frozenSplit: return "frozenSplit";
    }
  }
}

enum PivotAreaType { none, normal, data, all, origin, button, topEnd }

extension PivotAreaTypeExt on PivotAreaType {
  static PivotAreaType? parse(String str){
    switch(str){
      case "none": return PivotAreaType.none;
      case "normal": return PivotAreaType.normal;
      case "data": return PivotAreaType.data;
      case "all": return PivotAreaType.all;
      case "origin": return PivotAreaType.origin;
      case "button": return PivotAreaType.button;
      case "topEnd": return PivotAreaType.topEnd;
    }
  }

  String get value {
    switch(this){
      case PivotAreaType.none: return "none";
      case PivotAreaType.normal: return "normal";
      case PivotAreaType.data: return "data";
      case PivotAreaType.all: return "all";
      case PivotAreaType.origin: return "origin";
      case PivotAreaType.button: return "button";
      case PivotAreaType.topEnd: return "topEnd";
    }
  }
}

/// CT_FontName
class FontNameProperty with XmlNodeWrapper{
  ExcelFile file;

  FontNameProperty(this.file,XmlNode node){
    this.node = node;
  }

  String get val => node.getAttribute("*val")!;
}

/// CT_FontSize
class FontSizeProperty with XmlNodeWrapper{
  ExcelFile file;

  FontSizeProperty(this.file,XmlNode node){
    this.node = node;
  }

  double get val => node.getAttribute("*val").asDoubleOrNull()!;
}

enum UnderlineValues {
  single,
  double,
  singleAccounting,
  doubleAccounting,
  none
}

extension UnderlineValuesExt on UnderlineValues {
  static UnderlineValues? parse(String str) {
    switch (str) {
      case "single":
        return UnderlineValues.single;
      case "double":
        return UnderlineValues.double;
      case "singleAccounting":
        return UnderlineValues.singleAccounting;
      case "doubleAccounting":
        return UnderlineValues.doubleAccounting;
      case "none":
        return UnderlineValues.none;
    }
  }

  String get key {
    switch (this) {
      case UnderlineValues.single:
        return "single";
      case UnderlineValues.double:
        return "double";
      case UnderlineValues.singleAccounting:
        return "singleAccounting";
      case UnderlineValues.doubleAccounting:
        return "doubleAccounting";
      case UnderlineValues.none:
        return "none";
    }
  }
}

/// CT_UnderlineProperty
class UnderlineProperty with XmlNodeWrapper{
  ExcelFile file;

  UnderlineProperty(this.file,XmlNode node){
    this.node = node;
  }

  UnderlineValues? get val => UnderlineValuesExt.parse(node.getAttribute("*val") ?? "single");
}

enum VerticalAlignRun {
  baseline,
  superscript,
  subscript
}

extension VerticalAlignRunExt on VerticalAlignRun {
  static VerticalAlignRun? parse(String str) {
    switch (str) {
      case "baseline":
        return VerticalAlignRun.baseline;
      case "superscript":
        return VerticalAlignRun.superscript;
      case "subscript":
        return VerticalAlignRun.subscript;
    }
  }

  String get key {
    switch (this) {
      case VerticalAlignRun.baseline:
        return "baseline";
      case VerticalAlignRun.superscript:
        return "superscript";
      case VerticalAlignRun.subscript:
        return "subscript";
    }
  }
}


/// CT_VerticalAlignFontProperty
class VerticalAlignFontProperty with XmlNodeWrapper{
  ExcelFile file;

  VerticalAlignFontProperty(this.file,XmlNode node){
    this.node = node;
  }

  VerticalAlignRun get val => VerticalAlignRunExt.parse(node.getAttribute("*val") ?? "")!;
}


/// CT_IntProperty
class IntProperty with XmlNodeWrapper{
  ExcelFile file;

  IntProperty(this.file,XmlNode node){
    this.node = node;
  }

  int get val => node.getAttribute("*val").asIntOrNull()!;
}

/// CT_BooleanProperty
class BooleanProperty with XmlNodeWrapper{
  ExcelFile file;

  BooleanProperty(this.file,XmlNode node){
    this.node = node;
  }

  bool get val => node.getAttribute("*val").asBoolOrTrue();
}

class ColorProperty with XmlNodeWrapper{
  ExcelFile file;

  ColorProperty(this.file,XmlNode node){
    this.node = node;
  }

  bool? get auto => node.getAttribute("*auto").asBoolOrNull();

  int? get indexed => node.getAttribute("*indexed").asIntOrNull();

  String? get rgb => node.getAttribute("*rgb");

  int? get theme => node.getAttribute("*theme").asIntOrNull();

  double? get tint => node.getAttribute("*tint").asDoubleOrDefault(0.0);
}

enum FontScheme {
  none,
  major,
  minor
}

extension FontSchemeExt on FontScheme {
  static FontScheme? parse(String str) {
    switch (str) {
      case "none":
        return FontScheme.none;
      case "major":
        return FontScheme.major;
      case "minor":
        return FontScheme.minor;
    }
  }

  String get key {
    switch (this) {
      case FontScheme.none:
        return "none";
      case FontScheme.major:
        return "major";
      case FontScheme.minor:
        return "minor";
    }
  }
}

/// CT_FontScheme
class FontSchemeProperty with XmlNodeWrapper{
  ExcelFile file;

  FontSchemeProperty(this.file,XmlNode node){
    this.node = node;
  }

  FontScheme get val => FontSchemeExt.parse(node.getAttribute("*val") ?? "")!;
}

/// ST_Xstring
class XStringProperty with XmlNodeWrapper{
  ExcelFile file;

  XStringProperty(this.file,XmlNode node){
    this.node = node;
  }

  String get val => node.innerXML;
}




/// ST_FontId
class FontIdProperty with XmlNodeWrapper{
  ExcelFile file;

  FontIdProperty(this.file,XmlNode node){
    this.node = node;
  }

  int get val => node.innerXML.asIntOrNull()!;
}

enum PhoneticType {
  halfwidthKatakana,
  fullwidthKatakana,
  Hiragana,
  noConversion
}

extension PhoneticTypeExt on PhoneticType {
  static PhoneticType? parse(String str) {
    switch (str) {
      case "halfwidthKatakana":
        return PhoneticType.halfwidthKatakana;
      case "fullwidthKatakana":
        return PhoneticType.fullwidthKatakana;
      case "Hiragana":
        return PhoneticType.Hiragana;
      case "noConversion":
        return PhoneticType.noConversion;
    }
  }

  String get key {
    switch (this) {
      case PhoneticType.halfwidthKatakana:
        return "halfwidthKatakana";
      case PhoneticType.fullwidthKatakana:
        return "fullwidthKatakana";
      case PhoneticType.Hiragana:
        return "Hiragana";
      case PhoneticType.noConversion:
        return "noConversion";
    }
  }
}

enum PhoneticAlignment {
  noControl,
  left,
  center,
  distributed
}

extension PhoneticAlignmentExt on PhoneticAlignment {
  static PhoneticAlignment? parse(String str) {
    switch (str) {
      case "noControl":
        return PhoneticAlignment.noControl;
      case "left":
        return PhoneticAlignment.left;
      case "center":
        return PhoneticAlignment.center;
      case "distributed":
        return PhoneticAlignment.distributed;
    }
  }

  String get key {
    switch (this) {
      case PhoneticAlignment.noControl:
        return "noControl";
      case PhoneticAlignment.left:
        return "left";
      case PhoneticAlignment.center:
        return "center";
      case PhoneticAlignment.distributed:
        return "distributed";
    }
  }
}


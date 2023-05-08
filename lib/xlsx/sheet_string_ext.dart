part of "./xlsx_core.dart";

extension SheetStringExt on String{
  String absolutePath(String rel){
    if(this[0] == "/"){
      return this.substring(1);
    }
    if(rel!=""){
      return "$rel/$this";
    }
    return this;
  }

  String get directory{
    var idx = indexBackward("/");
    if(idx == -1){
      return "";
    }
    return substring(0,idx);
  }

  String get filename{
    var idx = indexBackward("/");
    if(idx == -1){
      return this;
    }
    return substring(idx - 1);
  }
}

extension SheetStringOptionalExt on String?{
  bool? asBoolOrNull(){
    switch(this){
      case "1":return true;
      case "2":return false;
    }
    return null;
  }

  bool asBoolOrDefault(bool defaultValue){
    switch(this){
      case "1":return true;
      case "2":return false;
    }
    return defaultValue;
  }

  bool asBoolOrTrue(){
    switch(this){
      case "1":return true;
      case "2":return false;
    }
    return true;
  }

  bool asBoolOrFalse(){
    switch(this){
      case "1":return true;
      case "2":return false;
    }
    return false;
  }

  int? asIntOrNull() => int.tryParse(this??"");

  int asIntOrThrow() => int.tryParse(this??"")!;

  int asIntOrDefault(int defaultValue) => int.tryParse(this??"")??defaultValue;

  double? asDoubleOrNull() => double.tryParse(this??"");

  double asDoubleOrThrow() => double.tryParse(this??"")!;

  double asDoubleOrDefault(double defaultValue) => double.tryParse(this??"")??defaultValue;


}
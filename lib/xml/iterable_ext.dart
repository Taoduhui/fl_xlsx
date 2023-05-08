part of 'xml_core.dart';

extension SheetIterableExtension<T> on Iterable<T>{
  T? firstWhereOrNull(bool Function(T element) test){
    for(var i in this){
      if(test(i)){
        return i;
      }
    }
    return null;
  }
}
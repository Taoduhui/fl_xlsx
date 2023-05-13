// import 'dart:io';

import 'package:fl_xlsx/fl_xlsx.dart';
// import 'package:flutter_test/flutter_test.dart';



Future<void> main() async {

  var ref = CellRef("\$A\$1");
  for(int i=1;i<100;i++){
    ref.colIdx = i;
    print(ref.toString());
  }
  for(int i=1;i<100;i++){
    ref.rowIdx = i;
    print(ref.toString());
  }
  ref.isColFixed = false;
  print(ref.toString());
  ref.isColFixed = true;
  print(ref.toString());

  ref.isRowFixed = false;
  print(ref.toString());
  ref.isRowFixed = true;
  print(ref.toString());
  // var file = ExcelFile
  //     .fromBytes(await File("./files/MergeCell.xlsx").readAsBytes());
  //
  // var sheet1 = file.workbook.sheets.firstWhere((s) => s.name == "Sheet1").workSheet;
}

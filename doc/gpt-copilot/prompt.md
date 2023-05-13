```xml
<xsd:complexType name="CT_SheetView">
    <xsd:attribute name="windowProtection" type="xsd:boolean" use="optional" default="false"/>
    <xsd:attribute name="showGridLines" type="xsd:boolean" use="optional" default="true"/>
    <xsd:attribute name="showRowColHeaders" type="xsd:boolean" use="optional"/>
    <xsd:attribute name="view" type="ST_SheetViewType" use="optional" default="normal"/>
    <xsd:attribute name="topLeftCell" type="ST_CellRef" use="optional"/>
    <xsd:attribute name="colorId" type="xsd:unsignedInt" use="optional" default="64"/>
    <xsd:attribute name="zoomScale" type="xsd:double" use="optional" default="0.0"/>
    <xsd:attribute name="zoomScaleNormal" type="xsd:unsignedInt" use="optional"/>
    <xsd:attribute name="zoomScaleSheetLayoutView" type="xsd:double" use="optional"/>
    <xsd:attribute name="zoomScalePageLayoutView" type="xsd:unsignedInt" use="required"/>
    <xsd:attribute name="workbookViewId" type="xsd:unsignedInt" use="required"/>
    
    <xsd:simpleType name="ST_SheetViewType">
        <xsd:restriction base="xsd:string">
            <xsd:enumeration value="normal"/>
            <xsd:enumeration value="pageBreakPreview"/>
            <xsd:enumeration value="pageLayout"/>
        </xsd:restriction>
    </xsd:simpleType>
    <xsd:simpleType name="ST_CellRef">
        <xsd:restriction base="xsd:string"/>
    </xsd:simpleType>
</xsd:complexType>
```
```dart
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
  }
  
  String get key{
    switch(this){
      case SheetViewType.normal : return "normal";
      case SheetViewType.pageBreakPreview : return "pageBreakPreview";
      case SheetViewType.pageLayout : return "pageLayout";
    }
  }
}

class SheetView with XmlNodeWrapper {
  SheetView(XmlNode node) {
    this.node = node;
  }

  bool get windowProtection => node.getAttribute("*windowProtection").asBoolOrFalse();
  bool get showGridLines => node.getAttribute("*showGridLines").asBoolOrTrue();
  bool get showRowColHeaders => node.getAttribute("*showRowColHeaders").asBoolOrNull();
  SheetViewType get viewType =>  SheetViewTypeExt.parse(node.getAttribute("*view")??"normal");
  String? get topLeftCell => node.getAttribute("*topLeftCell");
  int get colorId => node.getAttribute("*colorId").asIntOrDefault(64);
  double get zoomScale => node.getAttribute("*zoomScale").asDoubleOrDefault(0.0);
  int? get zoomScaleNormal => node.getAttribute("*zoomScaleNormal").asIntOrNull();
  double? get zoomScaleSheetLayoutView => node.getAttribute("*zoomScaleSheetLayoutView").asDoubleOrNull();
  double get zoomScalePageLayoutView => node.getAttribute("*zoomScalePageLayoutView").asDoubleOrNull()!;
  int get workbookViewId => node.getAttribute("*workbookViewId").asIntOrNull()!;
}
```
The XML code block and the Dart code block above are in a mutually interchangeable relationship. Next, you will act as a converter. When I input XML to you, you need to provide Dart code. When I input Dart code, you need to provide XML.
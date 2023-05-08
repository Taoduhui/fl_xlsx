part of "./xlsx_core.dart";

class SharedStringItemsIterator extends XmlChildNodeIterator<SharedStringItem>{
  ExcelFile file;
  SharedStringItemsIterator(this.file,super.parent);

  @override
  SharedStringItem build(XmlNode n) => SharedStringItem(file,n);

  @override
  bool selector(XmlNode n) => n.type == XmlElementType.Normal && n.name.removeNamespace() == "si";
}

class SharedStringItems extends Iterable<SharedStringItem> with XmlNodeWrapper{
  ExcelFile file;
  SharedStringItems(this.file,XmlNode node){
    this.node = node;
  }

  @override
  Iterator<SharedStringItem> get iterator => SharedStringItemsIterator(file,node);
}

class SharedStringTable with XmlNodeWrapper{
  ExcelFile file;

  late List<SharedStringItem> cachedSharedStrings;

  SharedStringItem at(int idx){
    if(file.config.enableSharedStringCache){
      return cachedSharedStrings[idx];
    }
    return items.elementAt(idx);
  }

  SharedStringTable(this.file,XmlNode node){
    this.node = node;
    if(file.config.enableSharedStringCache){
      cachedSharedStrings = items.toList();
    }
  }

  SharedStringItems get items => SharedStringItems(file,node);

  int? get count => node.getAttribute("count").asIntOrNull();

  int? get uniqueCount => node.getAttribute("uniqueCount").asIntOrNull();
}


/// CT_Rst
class SharedStringItem with XmlNodeWrapper{
  ExcelFile file;

  SharedStringItem(this.file,XmlNode node){
    this.node = node;
  }

  XStringProperty? get text{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "t");
    if(_node != null){
      return XStringProperty(file,_node);
    }
    return null;
  }

  PhoneticRuns get phoneticRuns => PhoneticRuns(file,node);

  RichTextRuns get richTextRuns => RichTextRuns(file,node);

  PhoneticRunProperties? get phoneticRunsProperties{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "phoneticPr");
    if(_node != null){
      return PhoneticRunProperties(file,_node);
    }
    return null;
  }

  int? get uniqueCount => node.getAttribute("uniqueCount").asIntOrNull();
}

class RichTextRunsIterator extends XmlChildNodeIterator<RichTextRun>{
  ExcelFile file;

  RichTextRunsIterator(this.file,super.parent);

  @override
  RichTextRun build(XmlNode n) => RichTextRun(file, n);

  @override
  bool selector(XmlNode n) => n.type == XmlElementType.Normal && n.name.removeNamespace() == "r";

}

class RichTextRuns extends Iterable<RichTextRun> with XmlNodeWrapper{
  ExcelFile file;

  RichTextRuns(this.file,XmlNode node){
    this.node = node;
  }

  @override
  Iterator<RichTextRun> get iterator => RichTextRunsIterator(file,node);

}

///CT_RElt
class RichTextRun with XmlNodeWrapper{
  ExcelFile file;

  RichTextRun(this.file,XmlNode node){
    this.node = node;
  }

  RichTextRunProperties? get runProperties{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "rPr");
    if(_node != null){
      return RichTextRunProperties(file,_node);
    }
    return null;
  }

  XStringProperty get text{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "t");
    if(_node != null){
      return XStringProperty(file,_node);
    }
    throw Exception("here should be a 't' element but not found: ${node.start}");
  }
}

/// CT_RPrElt
class RichTextRunProperties with XmlNodeWrapper{
  ExcelFile file;

  RichTextRunProperties(this.file,XmlNode node){
    this.node = node;
  }

  FontNameProperty? get fontName{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "rFont");
    if(_node != null){
      return FontNameProperty(file,_node);
    }
    return null;
  }

  IntProperty? get charset{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "charset");
    if(_node != null){
      return IntProperty(file,_node);
    }
    return null;
  }

  IntProperty? get family{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "family");
    if(_node != null){
      return IntProperty(file,_node);
    }
    return null;
  }

  BooleanProperty? get bold{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "b");
    if(_node != null){
      return BooleanProperty(file,_node);
    }
    return null;
  }

  BooleanProperty? get italic{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "i");
    if(_node != null){
      return BooleanProperty(file,_node);
    }
    return null;
  }

  BooleanProperty? get strike{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "strike");
    if(_node != null){
      return BooleanProperty(file,_node);
    }
    return null;
  }

  BooleanProperty? get outline{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "outline");
    if(_node != null){
      return BooleanProperty(file,_node);
    }
    return null;
  }

  BooleanProperty? get shadow{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "shadow");
    if(_node != null){
      return BooleanProperty(file,_node);
    }
    return null;
  }

  BooleanProperty? get condense{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "condense");
    if(_node != null){
      return BooleanProperty(file,_node);
    }
    return null;
  }

  BooleanProperty? get extend{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "extend");
    if(_node != null){
      return BooleanProperty(file,_node);
    }
    return null;
  }

  ColorProperty? get color{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "color");
    if(_node != null){
      return ColorProperty(file,_node);
    }
    return null;
  }

  FontSizeProperty? get fontSize{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "sz");
    if(_node != null){
      return FontSizeProperty(file,_node);
    }
    return null;
  }

  UnderlineProperty? get underline{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "u");
    if(_node != null){
      return UnderlineProperty(file,_node);
    }
    return null;
  }

  VerticalAlignFontProperty? get vertAlign{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "vertAlign");
    if(_node != null){
      return VerticalAlignFontProperty(file,_node);
    }
    return null;
  }

  FontSchemeProperty? get scheme{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "scheme");
    if(_node != null){
      return FontSchemeProperty(file,_node);
    }
    return null;
  }
}


class PhoneticRunsIterator extends XmlChildNodeIterator<PhoneticRun>{
  ExcelFile file;

  PhoneticRunsIterator(this.file,super.parent);

  @override
  PhoneticRun build(XmlNode n) => PhoneticRun(file, n);

  @override
  bool selector(XmlNode n) => n.type == XmlElementType.Normal && n.name.removeNamespace() == "rPh";

}

class PhoneticRuns extends Iterable<PhoneticRun> with XmlNodeWrapper{
  ExcelFile file;

  PhoneticRuns(this.file,XmlNode node){
    this.node = node;
  }

  @override
  Iterator<PhoneticRun> get iterator => PhoneticRunsIterator(file,node);

}

/// CT_PhoneticRun
class PhoneticRun with XmlNodeWrapper{
  ExcelFile file;

  PhoneticRun(this.file,XmlNode node){
    this.node = node;
  }

  XStringProperty? get text{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "t");
    if(_node != null){
      return XStringProperty(file,_node);
    }
    return null;
  }

  int get starterBaselineAlignment => node.getAttribute("sb").asIntOrNull()!;
  int get endingBaselineAlignment => node.getAttribute("eb").asIntOrNull()!;
}

/// CT_PhoneticPr
class PhoneticRunProperties with XmlNodeWrapper{
  ExcelFile file;

  PhoneticRunProperties(this.file,XmlNode node){
    this.node = node;
  }

  FontIdProperty? get fontId{
    var _node = node.into(selector: (c)=>c.type == XmlElementType.Normal && c.name.removeNamespace() == "fontId");
    if(_node != null){
      return FontIdProperty(file,_node);
    }
    return null;
  }

  PhoneticType get type => PhoneticTypeExt.parse(node.getAttribute("type") ?? "fullwidthKatakana")!;

  PhoneticAlignment get alignment => PhoneticAlignmentExt.parse(node.getAttribute("alignment") ?? "left")!;
}


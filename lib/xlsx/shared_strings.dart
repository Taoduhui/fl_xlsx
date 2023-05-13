part of "./xlsx_core.dart";

class SharedStringItemsIterator extends XmlChildNodeIterator<SharedStringItem>{
  ExcelFile file;
  SharedStringItemsIterator(this.file,super.parent);

  @override
  SharedStringItem build(XmlNode n) => SharedStringItem(file,n);

  @override
  bool selector(XmlNode n) => n.type == XmlElementType.start && n.name.removeNamespace() == "si";
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
    if(file._config.enableSharedStringCache){
      return cachedSharedStrings[idx];
    }
    return items.elementAt(idx);
  }

  SharedStringTable(this.file,XmlNode node){
    this.node = node;
    if(file._config.enableSharedStringCache){
      cachedSharedStrings = items.toList();
    }
  }

  SharedStringItems get items => SharedStringItems(file,node);

  int? get count => node.getAttribute("*count").asIntOrNull();

  int? get uniqueCount => node.getAttribute("*uniqueCount").asIntOrNull();
}


/// CT_Rst
class SharedStringItem with XmlNodeWrapper{
  ExcelFile file;

  SharedStringItem(this.file,XmlNode node){
    this.node = node;
  }

  XStringProperty? get text{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "t");
    if(cnode != null){
      return XStringProperty(file,cnode);
    }
    return null;
  }

  PhoneticRuns get phoneticRuns => PhoneticRuns(file,node);

  RichTextRuns get richTextRuns => RichTextRuns(file,node);

  PhoneticRunProperties? get phoneticRunsProperties{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "phoneticPr");
    if(cnode != null){
      return PhoneticRunProperties(file,cnode);
    }
    return null;
  }

  int? get uniqueCount => node.getAttribute("*uniqueCount").asIntOrNull();
}

class RichTextRunsIterator extends XmlChildNodeIterator<RichTextRun>{
  ExcelFile file;

  RichTextRunsIterator(this.file,super.parent);

  @override
  RichTextRun build(XmlNode n) => RichTextRun(file, n);

  @override
  bool selector(XmlNode n) => n.type == XmlElementType.start && n.name.removeNamespace() == "r";

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
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "rPr");
    if(cnode != null){
      return RichTextRunProperties(file,cnode);
    }
    return null;
  }

  XStringProperty get text{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "t");
    if(cnode != null){
      return XStringProperty(file,cnode);
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
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "rFont");
    if(cnode != null){
      return FontNameProperty(file,cnode);
    }
    return null;
  }

  IntProperty? get charset{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "charset");
    if(cnode != null){
      return IntProperty(file,cnode);
    }
    return null;
  }

  IntProperty? get family{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "family");
    if(cnode != null){
      return IntProperty(file,cnode);
    }
    return null;
  }

  BooleanProperty? get bold{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "b");
    if(cnode != null){
      return BooleanProperty(file,cnode);
    }
    return null;
  }

  BooleanProperty? get italic{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "i");
    if(cnode != null){
      return BooleanProperty(file,cnode);
    }
    return null;
  }

  BooleanProperty? get strike{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "strike");
    if(cnode != null){
      return BooleanProperty(file,cnode);
    }
    return null;
  }

  BooleanProperty? get outline{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "outline");
    if(cnode != null){
      return BooleanProperty(file,cnode);
    }
    return null;
  }

  BooleanProperty? get shadow{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "shadow");
    if(cnode != null){
      return BooleanProperty(file,cnode);
    }
    return null;
  }

  BooleanProperty? get condense{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "condense");
    if(cnode != null){
      return BooleanProperty(file,cnode);
    }
    return null;
  }

  BooleanProperty? get extend{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "extend");
    if(cnode != null){
      return BooleanProperty(file,cnode);
    }
    return null;
  }

  ColorProperty? get color{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "color");
    if(cnode != null){
      return ColorProperty(file,cnode);
    }
    return null;
  }

  FontSizeProperty? get fontSize{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "sz");
    if(cnode != null){
      return FontSizeProperty(file,cnode);
    }
    return null;
  }

  UnderlineProperty? get underline{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "u");
    if(cnode != null){
      return UnderlineProperty(file,cnode);
    }
    return null;
  }

  VerticalAlignFontProperty? get vertAlign{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "vertAlign");
    if(cnode != null){
      return VerticalAlignFontProperty(file,cnode);
    }
    return null;
  }

  FontSchemeProperty? get scheme{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "scheme");
    if(cnode != null){
      return FontSchemeProperty(file,cnode);
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
  bool selector(XmlNode n) => n.type == XmlElementType.start && n.name.removeNamespace() == "rPh";

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
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "t");
    if(cnode != null){
      return XStringProperty(file,cnode);
    }
    return null;
  }

  int get starterBaselineAlignment => node.getAttribute("*sb").asIntOrNull()!;
  int get endingBaselineAlignment => node.getAttribute("*eb").asIntOrNull()!;
}

/// CT_PhoneticPr
class PhoneticRunProperties with XmlNodeWrapper{
  ExcelFile file;

  PhoneticRunProperties(this.file,XmlNode node){
    this.node = node;
  }

  FontIdProperty? get fontId{
    var cnode = node.into(selector: (c)=>c.type == XmlElementType.start && c.name.removeNamespace() == "fontId");
    if(cnode != null){
      return FontIdProperty(file,cnode);
    }
    return null;
  }

  PhoneticType get type => PhoneticTypeExt.parse(node.getAttribute("*type") ?? "fullwidthKatakana")!;

  PhoneticAlignment get alignment => PhoneticAlignmentExt.parse(node.getAttribute("*alignment") ?? "left")!;
}


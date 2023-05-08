part of 'xml_core.dart';

enum XmlElementType{
  Normal,
  ProcessingInstruction,
  Comment,
  CDATA,
  DTD,
  End,
  Unknown
}


class XmlNodeDetail{
  late int _beginElementStart;
  late int _beginElementEnd;
  late int _endElementStart;
  late int _endElementEnd;
  late XmlElementType _type;
}

class XmlAttribute{
  late int  keyIdx;
  late int  equalIdx;
  late int  quotaIdx;
  late int  quotaEndIdx;
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

class XmlNode{
  XmlDocument document;
  int start;
  bool removed = false;
  bool mounted = false;
  XmlNodeDetail? _detail;
  XmlNodeDetail get detail{
    _detail ??= _parseDetail(document,start);
    return _detail!;
  }

  XmlElementType get type => _detail?._type ?? document.raw.getXmlElementType(start);

  String get beginElement => document.raw.substring(detail._beginElementStart,detail._beginElementEnd + 1);

  String get endElement => document.raw.substring(detail._endElementStart,detail._endElementEnd + 1);

  void mount(){
    if(!mounted){
      mounted = true;
      document.mountedNodes.add(this);
    }
  }

  void unmount(){
    mounted = false;
  }

  bool _update(int modifiedStart,int modifiedLength){
    if(modifiedStart>start){
      if(_detail != null){
        if(modifiedStart <= _detail!._endElementEnd){
          _detail!._endElementEnd += modifiedLength;
        }
        if(modifiedStart <= _detail!._endElementStart){
          _detail!._endElementStart += modifiedLength;
        }
        if(modifiedStart <= _detail!._beginElementEnd){
          _detail!._beginElementEnd += modifiedLength;
        }
        if(modifiedStart <= _detail!._beginElementStart){
          _detail!._beginElementStart += modifiedLength;
        }
      }
    }else if(modifiedStart < start || (modifiedStart == start && modifiedLength > 0)){
      start += modifiedLength;
      if(_detail != null){
        _detail!._beginElementStart += modifiedLength;
        _detail!._beginElementEnd += modifiedLength;
        _detail!._endElementStart += modifiedLength;
        _detail!._endElementEnd += modifiedLength;
      }
    }else if(modifiedStart == start && modifiedLength < 0){
      removed = true;
      unmount();
      return false;
    }
    return true;
  }

  XmlNode({required this.document,required this.start});

  bool addAttribute(String key,String value){
    if(detail._type == XmlElementType.Normal) {
      var newAttr = " $key=\"$value\"";
      var insertAt = document.raw.findNodeNameEnd(start);
      if(insertAt == -1){
        return false;
      }
      document.raw = document.raw.insertAfter(newAttr, insertAt);
      document._update(insertAt, newAttr.length);
      return true;
    }
    return false;
  }

  bool setAttribute(String key,String value){
    if(detail._type == XmlElementType.Normal) {
      var attr = getAttributeNode(key);
      if(attr != null){
        var modifiedLength = value.length - (attr.quotaEndIdx - attr.quotaIdx - 1);
        document.raw = document.raw.remove(attr.quotaIdx + 1, attr.quotaEndIdx);
        document.raw = document.raw.insertAfter(value, attr.quotaIdx);
        document._update(attr.quotaEndIdx + 1, modifiedLength);
        return true;
      }
      return false;
    }
    return false;
  }

  bool removeAttribute(String key){
    if(detail._type == XmlElementType.Normal) {
      var attr = getAttributeNode(key);
      if(attr != null){
        document.raw = document.raw.remove(attr.keyIdx - 1, attr.quotaEndIdx + 1);
        document._update(attr.quotaEndIdx + 1, -(attr.quotaEndIdx - attr.keyIdx + 2));
      }
      return false;
    }
    return false;
  }

  XmlAttribute? getAttributeNode(String key){
    if(detail._type == XmlElementType.Normal){
      var attr = XmlAttribute();
      attr.keyIdx = document.raw.indexInRange(" $key",start: detail._beginElementStart,end: detail._beginElementEnd);
      if(attr.keyIdx == -1){
        attr.keyIdx = document.raw.indexInRange(":$key",start: detail._beginElementStart,end: detail._beginElementEnd);
      }
      if(attr.keyIdx == -1){
        return null;
      }
      attr.keyIdx++;
      attr.equalIdx = document.raw.indexInRange("=",start: attr.keyIdx,end: detail._beginElementEnd);
      if(attr.equalIdx == -1){
        return null;
      }
      attr.quotaIdx = document.raw.indexInRange("\"",start: attr.equalIdx,end: detail._beginElementEnd);
      if(attr.quotaIdx == -1){
        return null;
      }
      attr.quotaEndIdx = document.raw.findEndQuotation(attr.quotaIdx);
      if(attr.quotaEndIdx == -1){
        return null;
      }
      if(attr.quotaEndIdx < attr.quotaIdx){
        return null;
      }
      return attr;
    }
    return null;
  }

  String? getAttribute(String key){
    if(detail._type == XmlElementType.Normal){
      var attr =  getAttributeNode(key);
      if(attr == null){
        return null;
      }
      return document.raw.substring(attr.quotaIdx + 1,attr.quotaEndIdx);
    }
    return null;
  }

  Map<String,String> getAttributes(){
    if(detail._type == XmlElementType.Normal){
      return document.raw.parseAttributes(detail._beginElementStart,detail._beginElementEnd);
    }
    return {};
  }

  String get name{
    if(detail._type == XmlElementType.Normal) {
      var nameEnd = document.raw.findNodeNameEnd(start);
      return document.raw.substring(start + 1, nameEnd + 1);
    }
    return "";
  }

  /// Find next inner node
  ///
  /// ```XML
  /// &lt;Person&gt;
  ///   &lt;Name&gt;&lt;/Name&gt;
  ///   &lt;Age&gt;&lt;/Age&gt;
  /// &lt;/Person&gt;
  /// ```
  /// such as Person.into() is Name
  XmlNode? into({XmlElementType? type,bool Function(XmlNode node)? selector}){
    var nodeStart = document.raw.findInnerNodeStart(start);
    if(nodeStart == -1){
      return null;
    }
    XmlNode? node = XmlNode(document: document, start: nodeStart) ;
    if((selector != null && !selector(node)) || (type != null && node.type != type)){
      node = node.next(type: type,selector: selector);
    }
    return node;
  }

  /// Find next parallel node
  ///
  /// ```XML
  /// &lt;Person&gt;
  ///   &lt;Name&gt;&lt;/Name&gt;
  ///   &lt;Age&gt;&lt;/Age&gt;
  /// &lt;/Person&gt;
  /// ```
  /// such as Name.next() is Age
  XmlNode? next({XmlElementType? type,bool Function(XmlNode node)? selector}){
    if(type != null || selector != null){
      XmlNode? node;
      while(true){
        node = (node??this).next();
        if(node != null && selector != null && !selector(node)){
          continue;
        }
        if(node != null && type != null && node.type != type){
          continue;
        }
        return node;
      }
    }
    var nodeStart = document.raw.findParallelNodeStart(start);
    if(nodeStart == -1){
      return null;
    }
    return XmlNode(document: document, start: nodeStart) ;
  }

  /// Get node inner XML String
  ///
  /// ```XML
  /// &lt;Person&gt;
  ///   &lt;Name&gt;&lt;/Name&gt;
  ///   &lt;Age&gt;&lt;/Age&gt;
  /// &lt;/Person&gt;
  /// ```
  /// such as Name.innerXML is
  /// \\t&lt;Name&gt;&lt;/Name&gt;\\t\\n&lt;Age&gt;&lt;/Age&gt;
  String get innerXML{
    if(detail._type == XmlElementType.Normal){
      return document.raw.substring(detail._beginElementEnd + 1,detail._endElementStart);
    }
    return "";
  }

  String get value{
    //TODO: parse special character
    return innerXML;
  }

  static XmlNodeDetail _parseDetail(XmlDocument document,int start){
    var detail = XmlNodeDetail();
    detail._type = document.raw.getXmlElementType(start);
    detail._beginElementStart = start;
    detail._endElementEnd = document.raw.findNodeEnd(start);
    detail._beginElementEnd = document.raw.findElementEnd(detail._beginElementStart,detail._type);
    detail._endElementStart = document.raw.findElementStart(detail._endElementEnd,detail._type);
    return detail;
  }
}
part of 'xml_core.dart';

extension XmlStringExtension on String{

  String removeNamespace(){
    var idx = indexOf(":");
    if(idx == -1){
      return this;
    }
    return substring(idx + 1);
  }

  int findEndQuotation(int start){
    var type = this[start];
    for(int i=start + 1;i<length;i++){
      if(this[i] == type){
        return i;
      }
    }
    throw(Exception("unquoted element, start:$start"));
  }

  int findDoubleQuotationBackward(int end){
    var type = this[end];
    for(int i=end - 1;i>=0;i--){
      if(this[i] == type){
        return i;
      }
    }
    throw(Exception("unquoted element, end:$end"));
  }

  int findNormalElementEnd(int start){
    for(int i=start + 1;i<length;i++){
      if(this[i] == "\""){
        i = findEndQuotation(i);
        continue;
      }
      if(this[i] == ">"){
        return i;
      }
    }
    throw(Exception("missing >, start:$start"));
  }

  int findNormalElementStart(int end){
    for(int i=end - 1;i>=0;i--){
      if(this[i] == "\""){
        i = findDoubleQuotationBackward(i);
        continue;
      }
      if(this[i] == "<"){
        return i;
      }
    }
    throw(Exception("missing <, end:$end"));
  }

  int indexBackward(String str,[int? end]){
    var _end = end ?? (length - 1);
    var strLastIdx = str.length - 1;
    for(var i = _end; i >= strLastIdx;i--){
      bool matched = true;
      for(var c = strLastIdx;c>=0;c--){
        if(this[i - (strLastIdx - c)] != str[c]){
          matched = false;
          break;
        }
      }
      if(matched){
        return i - strLastIdx;
      }
    }
    return -1;
  }

  int findCommentElementStart(int end){
    var start = indexBackward("<!--",end);
    if(start == -1){
      throw(Exception("missing Comment start element <!--, end:$end"));
    }
    return start;
  }

  int findCommentElementEnd(int start){
    var end = indexOf("-->",start);
    if(end == -1){
      throw(Exception("missing Comment end element -->, start:$start"));
    }
    return end + 2;
  }

  int findCDATAElementStart(int end){
    var start = indexOf("<![CDATA[",end);
    if(start == -1){
      throw(Exception("missing CDATA start element <![CDATA[, end:$end"));
    }
    return start;
  }

  int findCDATAElementEnd(int start){
    var end = indexOf("]]>",start);
    if(end == -1){
      throw(Exception("missing CDATA end element ]]>, start:$start"));
    }
    return end + 2;
  }

  int findProcessingInstructionElementStart(int end){
    var start = indexBackward("<?",end);
    if(start == -1){
      throw(Exception("missing processing instruction start element <?, end:$end"));
    }
    return start;
  }

  int findProcessingInstructionElementEnd(int start){
    var end = indexOf("?>",start);
    if(end == -1){
      throw(Exception("missing processing instruction end element ?>, start:$start"));
    }
    return end + 1;
  }

  int findDTDElementStart(int end){
    for(int i=end-1;i>=0;i--){
      if(this[i] == ">"){
        i = findNormalElementStart(i);
      }else if(this[i] == "<"){
        return i;
      }
    }
    throw(Exception("missing >, end:$end"));
  }

  int findDTDElementEnd(int start){
    for(int i=start + 2;i<length;i++){
      if(this[i] == "<"){
        i = findNormalElementEnd(i);
      }else if(this[i] == ">"){
        return i;
      }
    }
    throw(Exception("missing >, start:$start"));
  }

  int findInnerNodeStart(int start){
    var end = findNormalElementEnd(start);
    for(var i=end + 1;i<length;i++){
      if(this[i] == "<"){
        if(this[i+1] == "/"){
          return -1;
        }
        return i;
      }
    }
    return -1;
  }

  int findParallelNodeStart(int start){
    var end = findNodeEnd(start);
    for(var i=end + 1;i<length;i++){
      if(this[i] == "<"){
        if(this[i+1] == "/"){
          return -1;
        }
        return i;
      }
    }
    return -1;
  }

  int findElementStart(int end,[XmlElementType? type]){
    var _type = type ?? getXmlElementTypeBySuffix(end);
    switch(_type){
      case XmlElementType.Normal:return findNormalElementStart(end);
      case XmlElementType.ProcessingInstruction:return findProcessingInstructionElementStart(end);
      case XmlElementType.Comment:return findCommentElementStart(end);
      case XmlElementType.CDATA:return findCDATAElementStart(end);
      case XmlElementType.DTD:return findDTDElementStart(end);
      case XmlElementType.End:return findNormalElementStart(end);
      case XmlElementType.Unknown:return -1;
    }
  }

  int findElementEnd(int start,[XmlElementType? type]){
    var _type = type ?? getXmlElementType(start);
    switch(_type){
      case XmlElementType.Normal:return findNormalElementEnd(start);
      case XmlElementType.ProcessingInstruction:return findProcessingInstructionElementEnd(start);
      case XmlElementType.Comment:return findCommentElementEnd(start);
      case XmlElementType.CDATA:return findCDATAElementEnd(start);
      case XmlElementType.DTD:return findDTDElementEnd(start);
      case XmlElementType.End:return findNormalElementEnd(start);
      case XmlElementType.Unknown:return -1;
    }
  }

  int findNodeNameEnd(int start){
    var area = indexOf(RegExp("[ >]"),start);
    if(area == -1){
      return -1;
    }
    if(this[area] == ">" && this[area-1] == "/"){
      return area - 2;
    }
    return area - 1;
  }

  Map<String,String> parseAttributes(int start,int end){
    Map<String,String> map = {};
    start = findNodeNameEnd(start) + 1;
    end = end - 1;
    for(int i=start;i<=end;i++){
      var t = substring(i);
      if(this[i] == " " || this[i] == "/" || this[i] == ">"){
        continue;
      }
      var keyEndIdx = indexOf(RegExp("[ =]"),i);
      var key = substring(i,keyEndIdx);
      var quotaIdx = indexOf("\"",i);
      var quotaEndIdx = findEndQuotation(quotaIdx);
      var value = substring(quotaIdx + 1,quotaEndIdx);
      map[key] = value;
      i = quotaEndIdx;
    }
    return map;
  }

  int findNodeEnd(int start){
    int pair = 0;
    for(int i=start;i<length;i++){
      if(this[i] != "<"){
        continue;
      }
      var type = getXmlElementType(i);
      switch(type){
        case XmlElementType.ProcessingInstruction:{
          i = findProcessingInstructionElementEnd(i);
        };
        break;
        case XmlElementType.End:{
          i = findNormalElementEnd(i);
          pair--;
        }
        break;
        case XmlElementType.Comment:{
          i = findCommentElementEnd(i);
        }
        break;
        case XmlElementType.CDATA:{
          i = findCDATAElementEnd(i);
        }
        break;
        case XmlElementType.DTD:{
          i = findDTDElementEnd(i);
        }
        break;
        case XmlElementType.Normal:{
          i = findNormalElementEnd(i);
          if(this[i - 1] != "/"){
            pair++;
          }
        }
      }
      if(pair == 0){
        return i;
      }
    }
    return -1;
  }

  XmlElementType getXmlElementTypeBySuffix(int end){
    switch(this[end]){
      case ">":{
        switch(this[end - 1]){
          case "?":{
            // ?>: Processing instruction.
            return XmlElementType.ProcessingInstruction;
          }
          case "-":{
            // Probably --> for a comment.
            return XmlElementType.Comment;
          }
          case "]":{
            // Probably <![CDATA[...]]>
            if(this[end - 2] == "]"){
              return XmlElementType.CDATA;
            }
            return XmlElementType.DTD;
          }
          default:{
            var start = findNormalElementStart(end);
            var type = getXmlElementType(start);
            switch(type){
              case XmlElementType.Normal:return XmlElementType.Normal;
              case XmlElementType.End:return XmlElementType.End;
              default:return XmlElementType.Unknown;
            }
          }
        }
      }
    }
    return XmlElementType.Unknown;
  }

  XmlElementType getXmlElementType(int start){
    switch(this[start]){
      case "<":{
        switch(this[start + 1]){
          case "?":{
            // <?: Processing instruction.
            return XmlElementType.ProcessingInstruction;
          }
          case "/":{
            // </: End element
            return XmlElementType.End;
          }
          case "!":{
            // <!: Maybe comment, maybe CDATA.
            switch(this[start + 2]){
              case "-":{
                // Probably <!-- for a comment.
                return XmlElementType.Comment;
              }
              case "[":{
                // Probably <![CDATA[.
                return XmlElementType.CDATA;
              }
              default:{
                // Probably DTD, such as <!DOCTYPE ...>
                return XmlElementType.DTD;
              }
            }
          }
          default:{
            return XmlElementType.Normal;
          }
        }
      }
    }
    return XmlElementType.Unknown;
  }

  int indexInRange(String str,{int start = 0,int? end}){
    end ??= length - 1;
    for(int i=start;i<end;i++){
      bool matched = true;
      for(int c=0;c<str.length;c++){
        if(this[i + c] != str[c]){
          matched = false;
          break;
        }
      }
      if(matched){
        return i;
      }
    }
    return -1;
  }

  String insertAfter(String str,int index){
    var str1 = substring(0,index + 1);
    var str2 = substring(index + 1);
    return "$str1$str$str2";
  }

  String remove(int start,int end){
    var str1 = substring(0,start);
    var str2 = substring(end);
    return "$str1$str2";
  }
}
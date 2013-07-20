///<reference path="ActiveXObject.js">
!function() {

  var document = new Document
    , node     = new Node(document)
    , nodelist = new NodeList
    , namedmap = new NamedNodeMap

  function Node(owner) {
    ///<field name="attributes"      type="NamedNodeMap" >Contains the list of attributes for this node. Read-only.</field>
    ///<field name="baseName"        type="String"       >Returns the base name for the name qualified with the namespace. Read-only. For example, it returns "yyy" for the element &lt;xxx:yyy>.</field>
    ///<field name="childNodes"      type="NodeList"     >Contains a node list containing the children nodes. Read-only.</field>
    ///<field name="dataType"        type="Variant"      >Specifies the data type for this node. Read/write.</field>
    ///<field name="definition"      type="Variant"      >Returns the definition of the node in the document type definition (DTD) or schema. Read-only.</field>
    ///<field name="firstChild"      type="Node"         >Contains the first child of this node. Read-only.</field>
    ///<field name="lastChild"       type="Node"         >Returns the last child node. Read-only.</field>
    ///<field name="namespaceURI"    type="String"       >Returns the Uniform Resource Identifier (URI) for the namespace. Read-only. This refers to the "uuu" portion of the namespace declaration xmlns:nnn="uuu".</field>
    ///<field name="nextSibling"     type="Node"         >Contains the next sibling of the node in the parent's child list. Read-only.</field>
    ///<field name="nodeName"        type="String"       >Returns the qualified name for attribute, document type, element, entity, or notation nodes. Returns a fixed string for all other node types. Read-only.</field>
    ///<field name="nodeType"        type="Number"       >Specifies the XML Document Object Model (DOM) node type, which determines valid values and whether the node can have child nodes. Read-only.</field>
    ///<field name="nodeTypedValue"  type="Null"         >Contains this node's value expressed in its defined data type. Read/write.</field>
    ///<field name="nodeTypeString"  type="String"       >Returns the node type in string form. Read-only.</field>
    ///<field name="nodeValue"       type="Null"         >Contains the text associated with the node. Read/write.</field>
    ///<field name="ownerDocument"   type="Document"     >Returns the root of the document that contains this node. Read-only.</field>
    ///<field name="parentNode"      type="Node"         >Contains the parent node. Read-only.</field>
    ///<field name="parsed"          type="Boolean"      >Indicates the parsed status of the node and child nodes. Read-only.</field>
    ///<field name="prefix"          type="String"       >Returns the namespace prefix. Read-only. For example, for the element &lt;xxx:yyy>, it returns "xxx"</field>
    ///<field name="previousSibling" type="Node"         >Contains the previous sibling of the node in the parent's child list. Read-only.</field>
    ///<field name="specified"       type="Boolean"      >Indicates whether the node (usually an attribute) is explicitly specified or derived from a default value in the DTD or schema. Read-only.</field>
    ///<field name="text"            type="String"       >Represents the text content of the node or the concatenated text representing the node and its descendants. Read/write.</field>
    ///<field name="xml"             type="String"       >Contains the XML representation of the node and all its descendants. Read-only</field>
    this.childNodes      = nodelist
    this.firstChild      = 
    this.lastChild       = 
    this.nextSibling     = 
    this.parentNode      = 
    this.previousSibling = this
    this.ownerDocument   = owner || document
  }

  function Document() {
    this.childNodes      = nodelist
    this.documentElement = 
    this.firstChild      = 
    this.lastChild       = 
    this.nextSibling     = 
    this.parentNode      = 
    this.previousSibling = new Node(this)
    this.ownerDocument   = this
    ///<field name="attributes"      type="NamedNodeMap" >Contains the list of attributes for this node. Read-only.</field>
    ///<field name="childNodes"      type="NodeList"     >Contains a node list containing the children nodes. Read-only.</field>
    ///<field name="firstChild"      type="Node"         >Contains the first child of this node. Read-only.</field>
    ///<field name="lastChild"       type="Node"         >Returns the last child node. Read-only.</field>
    ///<field name="nextSibling"     type="Node"         >Contains the next sibling of the node in the parent's child list. Read-only.</field>
    ///<field name="parentNode"      type="Node"         >Contains the parent node. Read-only.</field>
    ///<field name="previousSibling" type="Node"         >Contains the previous sibling of the node in the parent's child list. Read-only.</field>

    ///<field name="async"              type="Boolean"        >Specifies if asynchronous download is permitted. Read/write.</field>
    ///<field name="doctype"            type="Document"       >Contains the document type node that specifies the DTD for this document. Read-only.</field>
    ///<field name="documentElement"    type="Node"           >Contains the root element of the document. Read/write.</field>
    ///<field name="implementation"     type="Implementation" >Contains the IXMLDOMImplementation object for the document. Read-only.</field>
    ///<field name="ondataavailable"    type="Function"       >Specifies the event handler for the ondataavailable event. Write-only.</field>
    ///<field name="onreadystatechange" type="Function"       >Specifies the event handler to be called when the readyState property changes. Write-only.</field>
    ///<field name="ontransformnode"    type="Function"       >Specifies the event handler for the ontransformnode event. Write-only.</field>
    ///<field name="parseError"         type="ParseError"     >Returns an IXMLDOMParseError object that contains information about the last parsing error. Read-only.</field>
    ///<field name="preserveWhiteSpace" type="Boolean"        >Specifies the default white space handling. Read/write.</field>
    ///<field name="readyState"         type="Number"         >Indicates the current state of the XML document. Read-only.</field>
    ///<field name="resolveExternals"   type="Boolean"        >Indicates whether external definitions (resolvable namespaces, DTD external subsets, and external entity references) are to be resolved at parse time, independent of validation. Read/write.</field>
    ///<field name="url"                type="String"         >Returns the URL for the last loaded XML document. Read-only.</field>
    ///<field name="validateOnParse"    type="Boolean"        >Indicates whether the parser should validate this document. Read/write.</field>

    ///<field name="namespaces" type="SchemaCache"></field>
    ///<field name="schemas" type="SchemaCache"></field>

  }

  function Implementation() {
    this.hasFeature = function(feature, version) {
      ///<param name="feature" type="String">"XML", "DOM", and "MS-DOM"</param>
      ///<param name="version" type="String" mayBeNull="true">"1.0"</param>
    }
  }

  function NamedNodeMap() {
    ///<field name="length" type="Number">Indicates the number of items in the collection. Read-only.</field>
    this.getNamedItem        = function() {
        ///<summary>Retrieves the attribute with the specified name.</summary>
    }
    this.getQualifiedItem    = function() {
        ///<summary>Returns the attribute with the specified namespace and attribute name.</summary>
    }
    this.item                = function() {
      ///<summary>Allows random access to individual nodes within the collection.</summary>
      ///<returns type="Node" />
    }
    this.nextNode            = function() {
      ///<summary>Returns the next node in the collection.</summary>
      ///<returns type="Node" />
    }
    this.reset               = function() {
      ///<summary>Resets the iterator.</summary>
    }
    this.removeNamedItem     = function() {
        ///<summary>Removes an attribute from the collection.</summary>
    }
    this.removeQualifiedItem = function() {
        ///<summary>Removes the attribute with the specified namespace and attribute name.</summary>
    }
    this.setNamedItem        = function() {
        ///<summary>Adds the supplied node to the collection.</summary>
    }
  }

  function NodeList() {
    ///<field name="length" type="Number">Indicates the number of items in the collection. Read-only.</field>
    this.item = function() {
      ///<summary>Allows random access to individual nodes within the collection.</summary>
      ///<returns type="Node" />
    }
    this.nextNode = function() {
      ///<summary>Returns the next node in the collection.</summary>
      ///<returns type="Node" />
    }
    this.reset = function() {
      ///<summary>Resets the iterator.</summary>
    }
  }

  function ParseError() {
    ///<field name="errorCode" type="Number">Contains the error code of the last parse error. Read-only.</field>
    ///<field name="filepos"   type="Number">Contains the absolute file position where the error occurred. Read-only.</field>
    ///<field name="line"      type="Number">Specifies the line number that contains the error. Read-only.</field>
    ///<field name="linepos"   type="Number">Contains the character position within the line where the error occurred. Read-only.</field>
    ///<field name="reason"    type="String">Describes the reason for the error. Read-only.</field>
    ///<field name="srcText"   type="String">Returns the full text of the line containing the error. Read-only.</field>
    ///<field name="url"       type="String">Contains the URL of the XML document containing the last error. Read-only.</field>
  }

  function SchemaCache() {
    ///<field name="length"         type="Number">Returns the number of namespaces currently in a collection. Read-only.</field>
    ///<field name="namespaceURI"   type="String">Returns the namespace at the specified index.</field>
    ///<field name="validateOnLoad" type="Boolean">Indicates whether schemas will be compiled and validated when loaded into the schema cache.</field>
    this.add            = function() {
      ///<summary>Adds a new schema to the schema collection, and associates the given namespace URI with the specified schema.</summary>
    }
    this.addCollection  = function() {
      ///<summary>Adds schemas from another collection into the current collection, and replaces any schemas that collide on the same namespace URI.</summary>
    }
    this.get            = function() {
      ///<summary>Returns a read-only Document Object Model (DOM) node that contains the &lt;Schema&gt; element.</summary>
    }
    this.getDeclaration = function() {
      ///<summary>Returns an ISchemaItem object for a specified DOM node. The return value will include declaration information. This information typically references the schema to be applied when validating the contents of the specified DOM node.</summary>
    }
    this.getSchema      = function() {
      ///<summary>Returns an ISchema object. The schema contains the namespace URI specified in the namespaceURI parameter that is passed to this method. The ISchema interface can be used to further obtain information about the schema object that is returned.</summary>
    }
    this.remove         = function() {
      ///<summary>Removes the specified namespace from a collection.</summary>
    }
    this.validate       = function() {
      ///<summary>Performs run-time validation on the documents in the schema cache that have not been compiled and validated.</summary>
    }
  }

// Node                                                           //

  Node.prototype = { constructor : Node
  , appendChild           : function(newChild) {
      ///<summary>Appends a new child as the last child of the node.</summary>
      ///<returns type="Node" value="newChild"/>
    }
  , cloneNode             : function(deep) {
      ///<summary>Clones a new node.</summary>
      ///<param name="deep" type="Boolean">
      ///A flag that indicates whether to recursively clone all nodes that are descendants of this node.&#10;
      ///If True, creates a clone of the complete tree below this node.&#10;
      ///If False, clones this node and its attributes only.
      ///</param>
      ///<returns type="Node"/>
    }
  , hasChildNodes         : function() {
      ///<summary>Provides a fast way to determine whether a node has children.</summary>
      ///<returns type="Boolean"/>
    }
  , insertBefore          : function(newChild, refChild) {
      ///<summary>Inserts a child node to the left of the specified node or at the end of the list.</summary>
      ///<param name="newChild" type="Node"/>
      ///<param name="refChild" type="Node" mayBeNull="true">If Null, the newChild parameter is inserted at the end of the child list.</param>
      ///<returns type="Node" value="newChild"/>
    }
  , removeChild           : function(childNode) {
      ///<summary>Removes the specified child node from the list of children and returns it.</summary>
      ///<param name="childNode" type="Node"/>
      ///<returns type="Node" value="childNode"/>
    }
  , replaceChild          : function(newChild, oldChild) {
      ///<summary>Replaces the specified old child node with the supplied new child node.</summary>
      ///<param name="newChild" type="Node" mayBeNull="true">If Null, oldChild is removed without a replacement</param>
      ///<param name="oldChild" type="Node"/>
      ///<returns type="Node" value="oldChild"/>
    }
  , selectNodes           : function(expression) {
      ///<summary>Applies the specified pattern-matching operation to this node's context and returns the list of matching nodes as IXMLDOMNodeList.</summary>
      ///<param name="expression" type="String">XPath expression.</param>
      ///<returns type="NodeList"/>
    }
  , selectSingleNode      : function(expression) {
      ///<summary>Applies the specified pattern-matching operation to this node's context and returns the first matching node.</summary>
      ///<param name="expression" type="String">XPath expression.</param>
      ///<returns type="Node"/>
    }
  , transformNode         : function(stylesheet) {
      ///<summary>Processes this node and its children using the supplied XSL Transformations (XSLT) style sheet and returns the resulting transformation.</summary>
      ///<param name="stylesheet" type="Document">A valid XML document or DOM node that consists of XSLT elements that direct the transformation of this node.</param>
      ///<returns type="String"/>
    }
  , transformNodeToObject : function(stylesheet, output) {
      ///<signature>
      ///<summary>Processes this node and its children using the supplied XSLT style sheet and returns the resulting transformation in the supplied object.</summary>
      ///<param name="stylesheet" type="Document">A valid XML document or DOM node that consists of XSLT elements that direct the transformation of this node.</param>
      ///<param name="output" type="Document"/>
      ///<returns type="Document"/>
      ///</signature>
      ///<signature>
      ///<summary>Processes this node and its children using the supplied XSLT style sheet and returns the resulting transformation in the supplied object.</summary>
      ///<param name="stylesheet" type="Document">A valid XML document or DOM node that consists of XSLT elements that direct the transformation of this node.</param>
      ///<param name="output" type="Stream"/>
      ///<returns type="Stream"/>
      ///</signature>
      ///<summary>Processes this node and its children using the supplied XSLT style sheet and returns the resulting transformation in the supplied object.</summary>
      ///<param name="stylesheet" type="Document">A valid XML document or DOM node that consists of XSLT elements that direct the transformation of this node.</param>
      ///<param name="output">DOMDocument object, or a stream.</param>
      ///<returns value="output"/>
    }
  }

// Document                                                       //

  Document.prototype = { constructor : Document
  , abort                       : function() {
      ///<summary>Aborts an asynchronous download in progress.</summary>
    }
  , createAttribute             : function() {
      ///<summary>Creates a new attribute with the specified name.</summary>
    }
  , createCDATASection          : function() {
      ///<summary>Creates a CDATA section node that contains the supplied data.</summary>
    }
  , createComment               : function() {
      ///<summary>Creates a comment node that contains the supplied data.</summary>
    }
  , createDocumentFragment      : function() {
      ///<summary>Creates an empty IXMLDOMDocumentFragment object.</summary>
    }
  , createElement               : function() {
      ///<summary>Creates an element node using the specified name.</summary>
    }
  , createEntityReference       : function() {
      ///<summary>Creates a new EntityReference object.</summary>
    }
  , createNode                  : function() {
      ///<summary>Creates a node using the supplied type, name, and namespace.</summary>
    }
  , createProcessingInstruction : function() {
      ///<summary>Creates a processing instruction node that contains the supplied target and data.</summary>
    }
  , createTextNode              : function() {
      ///<summary>Creates a text node that contains the supplied data.</summary>
    }
  , getElementsByTagName        : function() {
      ///<summary>Returns a collection of elements that have the specified name.</summary>
    }
  , load                        : function() {
      ///<summary>Loads an XML document from the specified location.</summary>
    }
  , loadXML                     : function() {
      ///<summary>Loads an XML document using the supplied string.</summary>
    }
  , nodeFromID                  : function() {
      ///<summary>Returns the node that matches the ID attribute.</summary>
    }
  , save                        : function() {
      ///<summary>Saves an XML document to the specified location.</summary>
    }
  , getProperty                 : function(name) {
      ///<summary>Retrieves the value of one of the "second-level properties" that are set either by default or using the setProperty method.</summary>
      ///<param name="name" type="String">
      ///AllowDocumentFunction  True&#10;
      ///ForcedResync           True&#10;
      ///MaxXMLSize             0&#10;
      ///MultipleErrorMessages  False&#10;
      ///NewParser              False&#10;
      ///ProhibitDTD            False&#10;
      ///ResolveExternals       True&#10;
      ///SelectionLanguage      "XPath" / "XSLPattern"&#10;
      ///SelectionNamespaces    ""&#10;
      ///ServerHTTPRequest      False&#10;
      ///UseInlineSchema        True&#10;
      ///ValidateOnParse        True&#10;
      ///</param>
      return SECOND_LEVEL_PROPERTIES[name]
    }
  , setProperty                 : function(name, value) {
      ///<summary>This method is used to set "second-level properties" on the DOM object.</summary>
      ///<signature><param name="name" type="String">
      ///AllowDocumentFunction  True&#10;
      ///ForcedResync           True&#10;
      ///MultipleErrorMessages  False&#10;
      ///NewParser              False&#10;
      ///ProhibitDTD            False&#10;
      ///ResolveExternals       True&#10;
      ///ServerHTTPRequest      False&#10;
      ///UseInlineSchema        True&#10;
      ///ValidateOnParse        True&#10;
      ///</param><param name="value" type="Boolean"/></signature>
      ///<signature><param name="name" type="String">
      ///MaxXMLSize             0&#10;
      ///</param><param name="value" type="Number"/></signature>
      ///<signature><param name="name" type="String">
      ///SelectionLanguage      "XPath" / "XSLPattern"&#10;
      ///SelectionNamespaces    ""&#10;
      ///</param><param name="value" type="String"/></signature>
    }
  , validate                    : function() {
      ///<summary>Performs run-time validation on the currently loaded document using the currently loaded document type definition (DTD), schema, or schema collection.</summary>
      ///<returns type="ParseError"/>
    }
  , validateNode                : function() {
      ///<summary>Only in MSXML 5.0 or greater</summary>
    }
  , importNode                  : function() {
      ///<summary>only in MSXML 5.0 or greater</summary>
    }
  }

  var D = Document.prototype
    , N = Node.prototype
    , p

  for (p in N) if (!(p in D)) D[p] = N[p]

// ENUM                                                           //
!function() {

  var SECOND_LEVEL_PROPERTIES = {
      'AllowDocumentFunction' : true    // in MSXML 4.0 and earlier. False in MSXML 5.0 for Microsoft Office Applications and later. 3.0 SP4, 4.0 SP2 and later**
    , 'ForcedResync'          : true    // 3.0 SP3, 4.0 SP1, 5.0 and later
    , 'MaxXMLSize'            : 0       // 3.0 SP4, 4.0 SP2 and later***
    , 'MultipleErrorMessages' : false   // 5.0 and later
    , 'NewParser'             : false   // 4.0 and later
    , 'ProhibitDTD'           : false   // 3.0 SP4 QFE and later, 4.0 SP2 QFE and later***
    , 'ResolveExternals'      : true    // 5.0 and later*
    , 'SelectionLanguage'     : 'XPath' // 3.0 and later "XSLPattern" in 3.0; "XPath" in 4.0 and later
    , 'SelectionNamespaces'   : ''      // 3.0 and later
    , 'ServerHTTPRequest'     : false   // 3.0 and later
    , 'UseInlineSchema'       : true    // 5.0 and later
    , 'ValidateOnParse'       : true    // 5.0 and later*
  }

  'microsoft msxml msxml2'.replace(/\w+/g, function(m) {
    var i = 9
    while (i--) {
      ActiveXObject[m+'.freethreadeddomdocument'+(i?'.'+i+'.0':'')] = Document
    }
  })
}()
}()

var xml   = new ActiveXObject('msxml.freethreadeddomdocument.3.0')
  , first = xml.firstChild
  , next  = first.firstChild.firstChild.lastChild.childNodes.nextNode()

xml

!function() {/*

    ///<summary></summary>
    ///<param name="" type=""></param>
    ///<returns type=""/>

<param>
  http://msdn.microsoft.com/en-us/library/hh542720.aspx

  // 2.x
  Microsoft.XMLDOM             , MSXML.DOMDocument
  Microsoft.FreeThreadedXMLDOM , MSXML.FreeThreadedDOMDocument
  Microsoft.XMLDSO
  Microsoft.XMLHTTP

  //.3.0
  MSXML2.DOMDocument
  MSXML2.FreeThreadedDOMDocument
  MSXML2.DSOControl
  MSXML2.XMLHTTP
  MSXML2.XMLSchemaCache
  MSXML2.XSLTemplate

  //.4.0
  Msxml2.DOMDocument
  Msxml2.DSOControl
  Msxml2.FreeThreadedDOMDocument
  Msxml2.MXHTMLWriter
  Msxml2. MXNamespaceManager
  Msxml2.MXXMLWriter
  Msxml2.SAXAttributes
  Msxml2.SAXXMLReader
  Msxml2.ServerXMLHTTP
  Msxml2.XMLHTTP
  Msxml2.XMLSchemaCache
  Msxml2.XSLTemplate

  //.5.0
  Msxml2.DOMDocument
  Msxml2.DSOControl
  Msxml2.FreeThreadedDOMDocument
  Msxml2.MXDigitalSignature
  Msxml2.MXHTMLWriter
  Msxml2.MXNamespaceManager
  Msxml2.MXXMLWriter
  Msxml2.SAXAttributes
  Msxml2.SAXXMLReader
  Msxml2.ServerXMLHTTP
  Msxml2.XMLHTTP
  Msxml2.XMLSchemaCache
  Msxml2.XSLTemplate

*/
}()
/// <reference path="ActiveXObject.js">
+function() {

  var document = new Document
  var node     = new Node(document)
  var nodelist = new NodeList
  var namedmap = new NamedNodeMap

  function Node(owner) {
    /// <field name="attributes"      type="NamedNodeMap" >Contains the list of attributes for this node. Read-only.</field>
    /// <field name="baseName"        type="String"       >Returns the base name for the name qualified with the namespace. Read-only. For example, it returns "yyy" for the element &lt;xxx:yyy>.</field>
    /// <field name="childNodes"      type="NodeList"     >Contains a node list containing the children nodes. Read-only.</field>
    /// <field name="dataType"        type="Variant"      >Specifies the data type for this node. Read/write.</field>
    /// <field name="definition"      type="Variant"      >Returns the definition of the node in the document type definition (DTD) or schema. Read-only.</field>
    /// <field name="firstChild"      type="Node"         >Contains the first child of this node. Read-only.</field>
    /// <field name="lastChild"       type="Node"         >Returns the last child node. Read-only.</field>
    /// <field name="namespaceURI"    type="String"       >Returns the Uniform Resource Identifier (URI) for the namespace. Read-only. This refers to the "uuu" portion of the namespace declaration xmlns:nnn="uuu".</field>
    /// <field name="nextSibling"     type="Node"         >Contains the next sibling of the node in the parent's child list. Read-only.</field>
    /// <field name="nodeName"        type="String"       >Returns the qualified name for attribute, document type, element, entity, or notation nodes. Returns a fixed string for all other node types. Read-only.</field>
    /// <field name="nodeType"        type="Number"       >Specifies the XML Document Object Model (DOM) node type, which determines valid values and whether the node can have child nodes. Read-only.</field>
    /// <field name="nodeTypedValue"  type="Null"         >Contains this node's value expressed in its defined data type. Read/write.</field>
    /// <field name="nodeTypeString"  type="String"       >Returns the node type in string form. Read-only.</field>
    /// <field name="nodeValue"       type="Null"         >Contains the text associated with the node. Read/write.</field>
    /// <field name="ownerDocument"   type="Document"     >Returns the root of the document that contains this node. Read-only.</field>
    /// <field name="parentNode"      type="Node"         >Contains the parent node. Read-only.</field>
    /// <field name="parsed"          type="Boolean"      >Indicates the parsed status of the node and child nodes. Read-only.</field>
    /// <field name="prefix"          type="String"       >Returns the namespace prefix. Read-only. For example, for the element &lt;xxx:yyy>, it returns "xxx"</field>
    /// <field name="previousSibling" type="Node"         >Contains the previous sibling of the node in the parent's child list. Read-only.</field>
    /// <field name="specified"       type="Boolean"      >Indicates whether the node (usually an attribute) is explicitly specified or derived from a default value in the DTD or schema. Read-only.</field>
    /// <field name="text"            type="String"       >Represents the text content of the node or the concatenated text representing the node and its descendants. Read/write.</field>
    /// <field name="xml"             type="String"       >Contains the XML representation of the node and all its descendants. Read-only</field>
    this.childNodes      = nodelist
    this.firstChild      = 
    this.lastChild       = 
    this.nextSibling     = 
    this.parentNode      = 
    this.previousSibling = this
    this.ownerDocument   = owner || document
  }

  function Document() {
    /// <field name="attributes"      type="NamedNodeMap" >Contains the list of attributes for this node. Read-only.</field>
    /// <field name="childNodes"      type="NodeList"     >Contains a node list containing the children nodes. Read-only.</field>
    /// <field name="firstChild"      type="Node"         >Contains the first child of this node. Read-only.</field>
    /// <field name="lastChild"       type="Node"         >Returns the last child node. Read-only.</field>
    /// <field name="nextSibling"     type="Node"         >Contains the next sibling of the node in the parent's child list. Read-only.</field>
    /// <field name="parentNode"      type="Node"         >Contains the parent node. Read-only.</field>
    /// <field name="previousSibling" type="Node"         >Contains the previous sibling of the node in the parent's child list. Read-only.</field>

    /// <field name="async"              type="Boolean"        >Specifies if asynchronous download is permitted. Read/write.</field>
    /// <field name="doctype"            type="Document"       >Contains the document type node that specifies the DTD for this document. Read-only.</field>
    /// <field name="documentElement"    type="Node"           >Contains the root element of the document. Read/write.</field>
    /// <field name="implementation"     type="Implementation" >Contains the IXMLDOMImplementation object for the document. Read-only.</field>
    /// <field name="ondataavailable"    type="Function"       >Specifies the event handler for the ondataavailable event. Write-only.</field>
    /// <field name="onreadystatechange" type="Function"       >Specifies the event handler to be called when the readyState property changes. Write-only.</field>
    /// <field name="ontransformnode"    type="Function"       >Specifies the event handler for the ontransformnode event. Write-only.</field>
    /// <field name="parseError"         type="ParseError"     >Returns an IXMLDOMParseError object that contains information about the last parsing error. Read-only.</field>
    /// <field name="preserveWhiteSpace" type="Boolean"        >Specifies the default white space handling. Read/write.</field>
    /// <field name="readyState"         type="Number"         >Indicates the current state of the XML document. Read-only.</field>
    /// <field name="resolveExternals"   type="Boolean"        >Indicates whether external definitions (resolvable namespaces, DTD external subsets, and external entity references) are to be resolved at parse time, independent of validation. Read/write.</field>
    /// <field name="url"                type="String"         >Returns the URL for the last loaded XML document. Read-only.</field>
    /// <field name="validateOnParse"    type="Boolean"        >Indicates whether the parser should validate this document. Read/write.</field>

    /// <field name="namespaces" type="XMLSchemaCache"></field>
    /// <field name="schemas" type="XMLSchemaCache"></field>

    this.childNodes      = nodelist
    this.documentElement = 
    this.firstChild      = 
    this.lastChild       = 
    this.nextSibling     = 
    this.parentNode      = 
    this.previousSibling = new Node(this)
    this.ownerDocument   = this

  }

  function Implementation() {
    this.hasFeature = function(feature, version) {
      /// <param name="feature" type="String">"XML", "DOM", and "MS-DOM"</param>
      /// <param name="version" type="String" mayBeNull="true">"1.0"</param>
    }
  }

  function NamedNodeMap() {
    /// <field name="length" type="Number">Indicates the number of items in the collection. Read-only.</field>
    this.getNamedItem        = function() {
        /// <summary>Retrieves the attribute with the specified name.</summary>
    }
    this.getQualifiedItem    = function() {
        /// <summary>Returns the attribute with the specified namespace and attribute name.</summary>
    }
    this.item                = function() {
      /// <summary>Allows random access to individual nodes within the collection.</summary>
      /// <returns type="Node" />
    }
    this.nextNode            = function() {
      /// <summary>Returns the next node in the collection.</summary>
      /// <returns type="Node" />
    }
    this.reset               = function() {
      /// <summary>Resets the iterator.</summary>
    }
    this.removeNamedItem     = function() {
        /// <summary>Removes an attribute from the collection.</summary>
    }
    this.removeQualifiedItem = function() {
        /// <summary>Removes the attribute with the specified namespace and attribute name.</summary>
    }
    this.setNamedItem        = function() {
        /// <summary>Adds the supplied node to the collection.</summary>
    }
  }

  function NodeList() {
    /// <field name="length" type="Number">Indicates the number of items in the collection. Read-only.</field>
    this.item     = function() {
      /// <summary>Allows random access to individual nodes within the collection.</summary>
      /// <returns type="Node" />
    }
    this.nextNode = function() {
      /// <summary>Returns the next node in the collection.</summary>
      /// <returns type="Node" />
    }
    this.reset    = function() {
      /// <summary>Resets the iterator.</summary>
    }
  }

  function ParseError() {
    /// <field name="errorCode" type="Number">Contains the error code of the last parse error. Read-only.</field>
    /// <field name="filepos"   type="Number">Contains the absolute file position where the error occurred. Read-only.</field>
    /// <field name="line"      type="Number">Specifies the line number that contains the error. Read-only.</field>
    /// <field name="linepos"   type="Number">Contains the character position within the line where the error occurred. Read-only.</field>
    /// <field name="reason"    type="String">Describes the reason for the error. Read-only.</field>
    /// <field name="srcText"   type="String">Returns the full text of the line containing the error. Read-only.</field>
    /// <field name="url"       type="String">Contains the URL of the XML document containing the last error. Read-only.</field>
  }

  function XMLHttp() {
    /// <field name="onreadystatechange" type="Function"/>
    /// <field name="readyState"         type="Number"/>
    /// <field name="responseBody"       type="Variant"/>
    /// <field name="responseStream"     type="Variant"/>
    /// <field name="responseText"       type="String"/>
    /// <field name="responseXML"        type="Document"/>
    /// <field name="status"             type="Number"/>
    /// <field name="statusText"         type="String"/>
    this.abort = function() {
      ///
    }
    this.getAllReponseHeaders = function() {
      /// <returns type="String"/>
    }
    this.getResponseHeader = function(header) {
      /// <param name="header" type="String"/>
      /// <returns type="String"/>
    }
    this.open = function(method, url, async, user, password) {
      /// <param name="method" type="String">GET , HEAD , POST</param>
      /// <param name="url" type="String"></param>
      /// <param name="async" type="Boolean"></param>
      /// <param name="user" type="String" optional="true"></param>
      /// <param name="password" type="String" optional="true"></param>
    }
    this.send = function(body) {
      /// <param name="body" type="Variant"></param>
    }
    this.setRequestHeader = function(header, value) {
      /// <param name="header" type="String"></param>
      /// <param name="value" type="String"></param>
    }
  }

  // XMLSchemaCache
  var XMLSchemaCache = function() { return XMLSchemaCache

    function XMLSchemaCache() {
      /// <field name="length"         type="Number"  >Returns the number of namespaces currently in a collection. Read-only.</field>
      /// <field name="namespaceURI"   type="String"  >Returns the namespace at the specified index.</field>
      /// <field name="validateOnLoad" type="Boolean" >Indicates whether schemas will be compiled and validated when loaded into the schema cache.</field>
      this.add            = function(namespaceURI, variant) {
        /// <summary>Adds a new schema to the schema collection, and associates the given namespace URI with the specified schema.</summary>
        /// <param name="namespaceURI" type="String">The namespace to associate with the specified schema. The empty string, "", will associate the schema with the empty namespace, xmlns="". &#10;&#10;This may be any string that can be used in an xmlns attribute, but it cannot contain entity references. The same white space normalization that occurs on the xmlns attribute also occurs on namespaceURI (that is, leading and trailing white space is trimmed, new lines are converted to spaces, and multiple adjacent white space characters are collapsed into one space).</param>
        /// <param name="variant">This specifies the schema to load. It will load it synchronously and with resolveExternals=false and validateOnParse=false. This parameter can also take any DOMDocument as an argument. &#10;&#10;This argument can be Null, which results in the removal of any schema for the specified namespaces. If the schema is an IXMLDOMNode, the entire document the node belongs to will be preserved.</param>
      }
      this.addCollection  = function(objXMLDOMSchemaCollection) {
        /// <summary>Adds schemas from another collection into the current collection, and replaces any schemas that collide on the same namespace URI.</summary>
        /// <param name="objXMLDOMSchemaCollection" type="XMLSchemaCache">The collection containing the schemas to add.</param>
      }
      this.get            = function(namespaceURI) {
        /// <summary>Returns a read-only Document Object Model (DOM) node that contains the &lt;Schema&gt; element.</summary>
        /// <param name="namespaceURI" type="String">The namespace URI associated with the schema to return. &#10;&#10;This can be any string that can be used in an xmlns attribute, but it cannot contain entity references. The same white space normalization that occurs on the xmlns attribute also occurs on this argument (that is, leading and trailing white space is trimmed, new lines are converted to spaces, and multiple adjacent white space characters are collapsed into one space).</param>
        /// <returns type="Document"/>
      }
      this.getDeclaration = function(oDOMNode) {
        /// <summary>Returns an ISchemaItem object for a specified DOM node. The return value will include declaration information. This information typically references the schema to be applied when validating the contents of the specified DOM node.</summary>
        /// <param name="oDOMNode" type="Node">An object. The DOM node for which a declaration object will be returned.</param>
        /// <returns type="SchemaItem"/>
      }
      this.getSchema      = function(strNamespaceURI) {
        /// <summary>Returns an ISchema object. The schema contains the namespace URI specified in the namespaceURI parameter that is passed to this method. The ISchema interface can be used to further obtain information about the schema object that is returned.</summary>
        /// <param name="strNamespaceURI" type="String">A string. The namespace of the schema to be retrieved from the schema cache.</param>
        /// <returns type="Schema"/>
      }
      this.remove         = function(namespaceURI) {
        /// <summary>Removes the specified namespace from a collection.</summary>
        /// <param name="namespaceURI" type="String">The namespace to remove from the collection. This can be any string that can be used in an xmlns attribute, but it cannot contain entity references. The same white space normalization that occurs on the xmlns attribute occurs on this argument (that is, leading and trailing white space is trimmed, new lines are converted to spaces, and multiple adjacent white space characters are collapsed into one space).</param>
      }
      this.validate       = function() {
        /// <summary>
        /// Performs run-time validation on the documents in the schema cache that have not been compiled and validated.
        /// If a schema has its validateOnLoad property set to true when it is loaded into the cache, the schema will be compiled and validated at that time.
        /// The validate method is used when validateOnLoad is set to false.
        /// Steps 2 and 3 are postponed until the validate method is called. &#10;&#10;
        /// The schema compilation steps are: &#10;&#10;
        /// 1. Check Syntax and resolve names.
        /// 2. Load and preprocess external schemas.
        /// 3. Compile schema items respecting W3Cs rules.
        /// </summary>
      }
    }

    function Schema() {
      /// <field name="attributeGroups" type="SchemaAttributeGroups" >Retrieves the collection of &lt;attributeGroup&gt; declarations. This collection contains objects that implement the ISchemaAttributeGroup interface.</field>
      /// <field name="attributes"      type="SchemaAttributes"      >Retrieves the collection of top-level &lt;attribute&gt; declarations. This collection contains objects that implement the ISchemaAttribute interface.</field>
      /// <field name="elements"        type="SchemaElements"        >Retrieves the collection of top-level &lt;element&gt; declarations. This collection contains objects that implement the ISchemaElement interface.</field>
      /// <field name="modelGroups"     type="SchemaModelGroups"     >Retrieves the collection of &lt;modelGroup&gt; declarations. This collection contains objects that implement the ISchemaModelGroup interface.</field>
      /// <field name="notations"       type="SchemaNotations"       >Retrieves the collection of notations. This collection contains objects that implement the ISchemaNotation interface.</field>
      /// <field name="schemaLocations" type="String"                >Retrieves the string collection of XML Schema URIs that are imported to or included in an XML Schema.</field>
      /// <field name="targetNamespace" type="String"                >Retrieves the value of the targetNamespace attribute for the XML Schema.</field>
      /// <field name="types"           type="SchemaTypes"           >Retrieves a collection of named simple-type and complex-type objects defined in the XML Schema.</field>
      /// <field name="version"         type="String"                >Retrieves the value of the version attribute of the XML Schema.</field>
    }

    function SchemaAny() {
      /// <field name="namespaces">Retrieves the collection of namespaces for the &lt;anyAttribute&gt; declaration.</field>
      /// <field name="processContents">Retrieves the processContents instructions for the &lt;anyAttribute&gt; declaration.</field>
    }

    function SchemaAttribute() {
      /// <field name="defaultValue" type="String"            >Retrieves the default value of the attribute.</field>
      /// <field name="fixedValue"   type="String"            >Retrieves the fixed value of the attribute.</field>
      /// <field name="isReference"  type="Boolean"           >Retrieves whether the attribute object is a reference to a top-level &lt;element&gt; declaration.</field>
      /// <field name="scope"        type="SchemaComplexType" >Retrieves the element definition type, either global or complex.</field>
      /// <field name="type"         type="SchemaType"        >Retrieves the type of the attribute.</field>
      /// <field name="use"          type="Number"            >Retrieves the use style for the attribute. &#10; SCHEMAUSE_OPTIONAL &#10; SCHEMAUSE_PROHIBITED &#10; SCHEMAUSE_REQUIRED</field>
    }

    function SchemaAttributes() {
      return SchemaItemCollection(SchemaAttribute)
    }

    function SchemaAttributeGroup() {
      /// <field name="anyAttribute" type="SchemaAny>Retrieves an ISchemaAny object with information about the &lt;anyAttribute&gt; element.</field>
      /// <field name="attributes" type="SchemaAttributes"/>
    }

    function SchemaAttributeGroups() {
      return SchemaItemCollection(SchemaAttributeGroup)
    }

    function SchemaComplexType() {
    
    }

    function SchemaElement() {}

    function SchemaElements() {
      return SchemaItemCollection(SchemaElement)
    }

    function SchemaModelGroup() {}

    function SchemaModelGroups() {
      return SchemaItemCollection(SchemaModelGroup)
    }

    function SchemaNotation() {}

    function SchemaNotations() {
      return SchemaItemCollection(SchemaNotation)
    }

    function SchemaItem() {
      /// <field name="id"                  type="String"               >Retrieves the id attribute of the schema item.</field>
      /// <field name="itemType"            type="Number"               >Retrieves the type of the schema item. An Enum of type SOMITEMTYPE.</field>
      /// <field name="name"                type="String"               >Retrieves the name of the item. The name is an NCName as defined by XML-Namespaces.</field>
      /// <field name="namespaceURI"        type="String"               >Retrieves the namespace URI of the item.</field>
      /// <field name="schema"              type="Schema"               >Retrieves the parent schema of the item.</field>
      /// <field name="unhandledAttributes" type="SchemaItemCollection" >Retrieves the collection of attributes that are not defined in the schema namespace.</field>
      this.writeAnnotation = function(annotationSink) {
        /// <summary>Writes an annotation to the XML Schema and to items in the XML Schema.</summary>
        /// <param name="annotationSink" type="Node">An object. A pointer to the annotation target object. The following interfaces can be sent: IXMLDOMDocument, IXMLDOMNode, IVBSAXContentHandler, and ISAXContentHandler.</param>
      }
      var SOMITEMTYPE = {
        SOMITEM_SCHEMA                      : 4096,  // 0x1000
        SOMITEM_ATTRIBUTE                   : 4097,  // 0x1001
        SOMITEM_ATTRIBUTEGROUP              : 4098,  // 0x1002
        SOMITEM_NOTATION                    : 4099,  // 0x1003
        SOMITEM_IDENTITYCONSTRAINT          : 4352,  // 0x1100
        SOMITEM_KEY                         : 4353,  // 0x1101
        SOMITEM_KEYREF                      : 4354,  // 0x1102
        SOMITEM_UNIQUE                      : 4355,  // 0x1103
        SOMITEM_ANYTYPE                     : 8192,  // 0x2000
        SOMITEM_DATATYPE                    : 8448,  // 0x2100
        SOMITEM_DATATYPE_ANYTYPE            : 8449,  // 0x2101
        SOMITEM_DATATYPE_ANYURI             : 8450,  // 0x2102
        SOMITEM_DATATYPE_BASE64BINARY       : 8451,  // 0x2103
        SOMITEM_DATATYPE_BOOLEAN            : 8452,  // 0x2104
        SOMITEM_DATATYPE_BYTE               : 8453,  // 0x2105
        SOMITEM_DATATYPE_DATE               : 88454, // 0x2106
        SOMITEM_DATATYPE_DATETIME           : 88455, // 0x2107
        SOMITEM_DATATYPE_DAY                : 88456, // 0x2108
        SOMITEM_DATATYPE_DECIMAL            : 88457, // 0x2109
        SOMITEM_DATATYPE_DOUBLE             : 88458, // 0x210A
        SOMITEM_DATATYPE_DURATION           : 88459, // 0x210B
        SOMITEM_DATATYPE_ENTITIES           : 88460, // 0x210C
        SOMITEM_DATATYPE_ENTITY             : 88461, // 0x210D
        SOMITEM_DATATYPE_FLOAT              : 88462, // 0x210E
        SOMITEM_DATATYPE_HEXBINARY          : 88463, // 0x210F
        SOMITEM_DATATYPE_ID                 : 88464, // 0x2110
        SOMITEM_DATATYPE_IDREF              : 8465,  // 0x2111
        SOMITEM_DATATYPE_IDREFS             : 8466,  // 0x2112
        SOMITEM_DATATYPE_INT                : 8467,  // 0x2113
        SOMITEM_DATATYPE_INTEGER            : 8468,  // 0x2114
        SOMITEM_DATATYPE_LANGUAGE           : 8469,  // 0x2115
        SOMITEM_DATATYPE_LONG               : 8470,  // 0x2116
        SOMITEM_DATATYPE_MONTH              : 8471,  // 0x2117
        SOMITEM_DATATYPE_MONTHDAY           : 8472,  // 0x2118
        SOMITEM_DATATYPE_NAME               : 8473,  // 0x2119
        SOMITEM_DATATYPE_NCNAME             : 8474,  // 0x211A
        SOMITEM_DATATYPE_NEGATIVEINTEGER    : 8475,  // 0x211B
        SOMITEM_DATATYPE_NMTOKEN            : 8476,  // 0x211C
        SOMITEM_DATATYPE_NMTOKENS           : 8477,  // 0x211D
        SOMITEM_DATATYPE_NONNEGATIVEINTEGER : 8478,  // 0x211E
        SOMITEM_DATATYPE_NONPOSITIVEINTEGER : 8479,  // 0x211F
        SOMITEM_DATATYPE_NORMALIZEDSTRING   : 8480,  // 0x2120
        SOMITEM_DATATYPE_NOTATION           : 8481,  // 0x2121
        SOMITEM_DATATYPE_POSITIVEINTEGER    : 8482,  // 0x2122
        SOMITEM_DATATYPE_QNAME              : 8483,  // 0x2123
        SOMITEM_DATATYPE_SHORT              : 8484,  // 0x2124
        SOMITEM_DATATYPE_STRING             : 8485,  // 0x2125
        SOMITEM_DATATYPE_TIME               : 8486,  // 0x2126
        SOMITEM_DATATYPE_TOKEN              : 8487,  // 0x2127
        SOMITEM_DATATYPE_UNSIGNEDBYTE       : 8488,  // 0x2128
        SOMITEM_DATATYPE_UNSIGNEDINT        : 8489,  // 0x2129
        SOMITEM_DATATYPE_UNSIGNEDLONG       : 8490,  // 0x212A
        SOMITEM_DATATYPE_UNSIGNEDSHORT      : 8491,  // 0x212B
        SOMITEM_DATATYPE_YEAR               : 8492,  // 0x212C
        SOMITEM_DATATYPE_YEARMONTH          : 8493,  // 0x212D
        SOMITEM_SIMPLETYPE                  : 8704,  // 0x2200
        SOMITEM_COMPLEXTYPE                 : 9216,  // 0x2400
        SOMITEM_PARTICLE                    : 16384, // 0x4000 // particle mask
        SOMITEM_ANY                         : 16385, // 0x4001
        SOMITEM_ANYATTRIBUTE                : 16386, // 0x4002
        SOMITEM_ELEMENT                     : 16387, // 0x4003
        SOMITEM_GROUP                       : 16640, // 0x4100 // group mask
        SOMITEM_ALL                         : 16641, // 0x4101
        SOMITEM_CHOICE                      : 16642, // 0x4102
        SOMITEM_SEQUENCE                    : 16643, // 0x4103
        SOMITEM_EMPTYPARTICLE               : 16644, // 0x4104
        SOMITEM_NULL                        : 2048,  // 0x0800 // null items
        SOMITEM_NULL_TYPE                   : 10240, // 0x2800
        SOMITEM_NULL_ANY                    : 18433, // 0x4801
        SOMITEM_NULL_ANYATTRIBUTE           : 18434, // 0x4802
        SOMITEM_NULL_ELEMENT                : 18435  // 0x4803
      }
    }

    function SchemaItemCollection(o) {
      if (!o) o = SchemaItem; o = new o()
      return {
        length: 0,
        item: function() {return o},
        itemByName: function(name) {return o},
        itemByQName: function(name, namespaceURI) {return o}
      }
    }

    function SchemaType() {
      /// <field name="baseTypes">Retrieves the collection of the base-types of the type.</field>
      /// <field name="derivedBy">Retrieves the method by which the type definition was derived.</field>
      /// <field name="enumeration">Retrieves the enumeration facet of the restriction.</field>
      /// <field name="final">Retrieves the final value of the type definition.</field>
      /// <field name="fractionDigits">Retrieves the fractionDigits facet of the restriction.</field>
      /// <field name="length">Retrieves the length facet of the restriction.</field>
      /// <field name="maxExclusive">Retrieves the maxExclusive facet of the restriction.</field>
      /// <field name="maxInclusive">Retrieves the maxInclusive facet of the restriction.</field>
      /// <field name="maxLength">Retrieves the maxLength facet of the restriction.</field>
      /// <field name="minExclusive">Retrieves the minExclusive facet of the restriction.</field>
      /// <field name="minInclusive">Retrieves the minInclusive facet of the restriction.</field>
      /// <field name="minLength">Retrieves the minLength facet of the restriction.</field>
      /// <field name="patterns">Retrieves a string collection that contains the pattern facets of the restriction.</field>
      /// <field name="totalDigits">Retrieves the totalDigits facet of the restriction.</field>
      /// <field name="variety">Retrieves the value of the variety attribute for the type definition.</field>
      /// <field name="whiteSpace">Retrieves the whitespace facet of the restriction.</field>
      this.isValid = function() {
        /// <summary>Checks whether the value that has been passed in is a valid instance of this type.</summary>
      }
    }

    function SchemaTypes() {
      return SchemaItemCollection(SchemaTypes)
    }

  }()

  function XSLTemplate() {
    this.stylesheet = node
    this.createProcessor = function() {
      return new XSLProcessor
    }
  }

  function XSLProcessor() {
    /// <field name="input"         type="Variant">XML input tree to transform</field>
    /// <field name="output"        type="Variant">custom stream object for transform output</field>
    /// <field name="ownerTemplate" type="XSLTemplate">template object used to create this processor object</field>
    /// <field name="readyState"    type="Number">current state of the processor</field>
    /// <field name="startMode"     type="String">starting XSL mode</field>
    /// <field name="startModeURI"  type="String">namespace of starting XSL mode</field>
    /// <field name="stylesheet"    type="Node">current stylesheet being used</field>
    this.addObject = function(object, namespaceURI) {
      /// <summary>pass object to stylesheet</summary>
      /// <param name="object" type="Object"/>
      /// <param name="namespaceURI" type="String" optional="true"/>
    }
    this.addParameter = function(baseName, parameter, namespaceURI) {
      /// <summary>set &lt;xsl:param&gt; values</summary>
      /// <param name="baseName" type="String"/>
      /// <param name="parameter" type="Variant"/>
      /// <param name="namespaceURI" type="String" optional="true"/>
    }
    this.reset = function() {
      /// <summary>reset state of processor and abort current transform</summary>
    }
    this.setStartMode = function(mode, namespaceURI) {
      /// <summary>set XSL mode and it's namespace</summary>
      /// <param name="mode" type="String"/>
      /// <param name="namespaceURI" type="String" optional="true"/>
    }
    this.transform = function() {
      /// <summary>start/resume the XSL transformation process</summary>
      /// <returns type="Boolean"/>
    }
  }

  function extendPrototypes() {

    var D = {
      abort                       : function() {
        /// <summary>Aborts an asynchronous download in progress.</summary>
      },
      createAttribute             : function() {
        /// <summary>Creates a new attribute with the specified name.</summary>
      },
      createCDATASection          : function() {
        /// <summary>Creates a CDATA section node that contains the supplied data.</summary>
      },
      createComment               : function() {
        /// <summary>Creates a comment node that contains the supplied data.</summary>
      },
      createDocumentFragment      : function() {
        /// <summary>Creates an empty IXMLDOMDocumentFragment object.</summary>
      },
      createElement               : function() {
        /// <summary>Creates an element node using the specified name.</summary>
      },
      createEntityReference       : function() {
        /// <summary>Creates a new EntityReference object.</summary>
      },
      createNode                  : function() {
        /// <summary>Creates a node using the supplied type, name, and namespace.</summary>
      },
      createProcessingInstruction : function() {
        /// <summary>Creates a processing instruction node that contains the supplied target and data.</summary>
      },
      createTextNode              : function() {
        /// <summary>Creates a text node that contains the supplied data.</summary>
      },
      getElementsByTagName        : function() {
        /// <summary>Returns a collection of elements that have the specified name.</summary>
      },
      load                        : function() {
        /// <summary>Loads an XML document from the specified location.</summary>
      },
      loadXML                     : function() {
        /// <summary>Loads an XML document using the supplied string.</summary>
      },
      nodeFromID                  : function() {
        /// <summary>Returns the node that matches the ID attribute.</summary>
      },
      save                        : function() {
        /// <summary>Saves an XML document to the specified location.</summary>
      },
      getProperty                 : function(name) {
        /// <summary>Retrieves the value of one of the "second-level properties" that are set either by default or using the setProperty method.</summary>
        /// <param name="name" type="String">
        /// AllowDocumentFunction  True                     &#10;
        /// ForcedResync           True                     &#10;
        /// MaxXMLSize             0                        &#10;
        /// MultipleErrorMessages  False                    &#10;
        /// NewParser              False                    &#10;
        /// ProhibitDTD            False                    &#10;
        /// ResolveExternals       True                     &#10;
        /// SelectionLanguage      "XPath" / "XSLPattern"   &#10;
        /// SelectionNamespaces    ""                       &#10;
        /// ServerHTTPRequest      False                    &#10;
        /// UseInlineSchema        True                     &#10;
        /// ValidateOnParse        True                     &#10;
        /// </param>
        return SECOND_LEVEL_PROPERTIES[name]
      },
      setProperty                 : function(name, value) {
        /// <summary>This method is used to set "second-level properties" on the DOM object.</summary>
        /// <signature><param name="name" type="String">
        /// AllowDocumentFunction  True  &#10;
        /// ForcedResync           True  &#10;
        /// MultipleErrorMessages  False &#10;
        /// NewParser              False &#10;
        /// ProhibitDTD            False &#10;
        /// ResolveExternals       True  &#10;
        /// ServerHTTPRequest      False &#10;
        /// UseInlineSchema        True  &#10;
        /// ValidateOnParse        True  &#10;
        /// </param><param name="value" type="Boolean"/></signature>
        /// <signature><param name="name" type="String">
        /// MaxXMLSize             0     &#10;
        /// </param><param name="value" type="Number"/></signature>
        /// <signature><param name="name" type="String">
        /// SelectionLanguage      "XPath" / "XSLPattern" &#10;
        /// SelectionNamespaces    ""                     &#10;
        /// </param><param name="value" type="String"/></signature>
      },
      validate                    : function() {
        /// <summary>Performs run-time validation on the currently loaded document using the currently loaded document type definition (DTD), schema, or schema collection.</summary>
        /// <returns type="ParseError"/>
      },
      validateNode                : function() {
        /// <summary>Only in MSXML 5.0 or greater</summary>
      },
      importNode                  : function() {
        /// <summary>only in MSXML 5.0 or greater</summary>
      }
    }

    var N = {
      appendChild           : function(newChild) {
        /// <summary>Appends a new child as the last child of the node.</summary>
        /// <returns type="Node" value="newChild"/>
      },
      cloneNode             : function(deep) {
        /// <summary>Clones a new node.</summary>
        /// <param name="deep" type="Boolean">
        /// A flag that indicates whether to recursively clone all nodes that are descendants of this node.&#10;
        /// If True, creates a clone of the complete tree below this node.&#10;
        /// If False, clones this node and its attributes only.
        /// </param>
        /// <returns type="Node"/>
      },
      hasChildNodes         : function() {
        /// <summary>Provides a fast way to determine whether a node has children.</summary>
        /// <returns type="Boolean"/>
      },
      insertBefore          : function(newChild, refChild) {
        /// <summary>Inserts a child node to the left of the specified node or at the end of the list.</summary>
        /// <param name="newChild" type="Node"/>
        /// <param name="refChild" type="Node" mayBeNull="true">If Null, the newChild parameter is inserted at the end of the child list.</param>
        /// <returns type="Node" value="newChild"/>
      },
      removeChild           : function(childNode) {
        /// <summary>Removes the specified child node from the list of children and returns it.</summary>
        /// <param name="childNode" type="Node"/>
        /// <returns type="Node" value="childNode"/>
      },
      replaceChild          : function(newChild, oldChild) {
        /// <summary>Replaces the specified old child node with the supplied new child node.</summary>
        /// <param name="newChild" type="Node" mayBeNull="true">If Null, oldChild is removed without a replacement</param>
        /// <param name="oldChild" type="Node"/>
        /// <returns type="Node" value="oldChild"/>
      },
      selectNodes           : function(expression) {
        /// <summary>Applies the specified pattern-matching operation to this node's context and returns the list of matching nodes as IXMLDOMNodeList.</summary>
        /// <param name="expression" type="String">XPath expression.</param>
        /// <returns type="NodeList"/>
      },
      selectSingleNode      : function(expression) {
        /// <summary>Applies the specified pattern-matching operation to this node's context and returns the first matching node.</summary>
        /// <param name="expression" type="String">XPath expression.</param>
        /// <returns type="Node"/>
      },
      transformNode         : function(stylesheet) {
        /// <summary>Processes this node and its children using the supplied XSL Transformations (XSLT) style sheet and returns the resulting transformation.</summary>
        /// <param name="stylesheet" type="Document">A valid XML document or DOM node that consists of XSLT elements that direct the transformation of this node.</param>
        /// <returns type="String"/>
      },
      transformNodeToObject : function(stylesheet, output) {
        /// <signature>
        /// <summary>Processes this node and its children using the supplied XSLT style sheet and returns the resulting transformation in the supplied object.</summary>
        /// <param name="stylesheet" type="Document">A valid XML document or DOM node that consists of XSLT elements that direct the transformation of this node.</param>
        /// <param name="output" type="Document"/>
        /// <returns type="Document"/>
        /// </signature>
        /// <signature>
        /// <summary>Processes this node and its children using the supplied XSLT style sheet and returns the resulting transformation in the supplied object.</summary>
        /// <param name="stylesheet" type="Document">A valid XML document or DOM node that consists of XSLT elements that direct the transformation of this node.</param>
        /// <param name="output" type="Stream"/>
        /// <returns type="Stream"/>
        /// </signature>
        /// <summary>Processes this node and its children using the supplied XSLT style sheet and returns the resulting transformation in the supplied object.</summary>
        /// <param name="stylesheet" type="Document">A valid XML document or DOM node that consists of XSLT elements that direct the transformation of this node.</param>
        /// <param name="output">DOMDocument object, or a stream.</param>
        /// <returns value="output"/>
      }
    }

    var p

    for (p in D) {
      Document.prototype[p] = D[p]
    }

    for (p in N) {
      Node.prototype[p] = N[p]
      if (!(p in D)) {
        Document.prototype[p] = N[p]
      }
    }

  }

  function connectInterfaces() {

    var SECOND_LEVEL_PROPERTIES = {
      AllowDocumentFunction : true,    // in MSXML 4.0 and earlier. False in MSXML 5.0 for Microsoft Office Applications and later. 3.0 SP4, 4.0 SP2 and later**
      ForcedResync          : true,    // 3.0 SP3, 4.0 SP1, 5.0 and later
      MaxXMLSize            : 0,       // 3.0 SP4, 4.0 SP2 and later***
      MultipleErrorMessages : false,   // 5.0 and later
      NewParser             : false,   // 4.0 and later
      ProhibitDTD           : false,   // 3.0 SP4 QFE and later, 4.0 SP2 QFE and later***
      ResolveExternals      : true,    // 5.0 and later*
      SelectionLanguage     : 'XPath', // 3.0 and later "XSLPattern" in 3.0; "XPath" in 4.0 and later
      SelectionNamespaces   : '',      // 3.0 and later
      ServerHTTPRequest     : false,   // 3.0 and later
      UseInlineSchema       : true,    // 5.0 and later
      ValidateOnParse       : true     // 5.0 and later*
    }

    'microsoft msxml msxml2'.replace(/\w+/g, function(w) {
      var i = 9
      while (i--) {
        v = (i ? '.'+ i +'.0' : '')
        ActiveXObject[w +'.domdocument'+ v] = Document;
        ActiveXObject[w +'.freethreadeddomdocument'+ v] = Document;
        ActiveXObject[w +'.xmlhttp'+ v] = XMLHttp;
        ActiveXObject[w +'.xmlschemacache'+ v] = XMLSchemaCache;
        ActiveXObject[w +'.xsltemplate'+ v] = XSLTemplate;
      }
    })

  }

  extendPrototypes()
  connectInterfaces()

}();

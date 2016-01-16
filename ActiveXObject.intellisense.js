var ActiveXObject = ActiveXObject || function (id) {
  /// <param name="id">
  /// ADODB.Stream                   &#10;
  /// Msxml2.DOMDocument             &#10;
  /// Msxml2.FreeThreadedDOMDocument &#10;
  /// Msxml2.XMLHTTP                 &#10;
  /// Msxml2.XMLSchemaCache          &#10;
  /// Msxml2.XSLTemplate             &#10;
  /// Scripting.FileSystemObject     &#10;
  /// Shell.Application              &#10;
  /// WScript.Network                &#10;
  /// WScript.Shell
  /// </param>
  return new ActiveXObject[String(arguments[0]).toLowerCase()]
};

+function() {

  ActiveXObject['adodb.stream'] = AdodbStream

  function AdodbStream() {
    /// <field name="CharSet" type="String">Sets or returns a value that specifies into which character set the contents are to be translated. This property is only used with text Stream objects (type is adTypeText)</field>
    /// <field name="EOS" type="Boolean">Returns whether the current position is at the end of the stream or not</field>
    /// <field name="LineSeparator" type="Number">Sets or returns the line separator character used in a text Stream object
    /// adCRLF -1 Carriage return line feed [Default] &#10;
    /// adLF   10 Line feed only                      &#10;
    /// adCR   13 Carriage return only                &#10;
    /// </field>
    /// <field name="Mode" type="Number">Sets or returns the available permissions for modifying data &#10;
    /// adModeUnknown        0 Permissions have not been set or cannot be determined.            &#10;
    /// adModeRead           1 Read-only.                                                        &#10;
    /// adModeWrite          2 Write-only.                                                       &#10;
    /// adModeReadWrite      3 Read/write.                                                       &#10;
    /// adModeShareDenyRead  4 Prevents others from opening a connection with read permissions.  &#10;
    /// adModeShareDenyWrite 8 Prevents others from opening a connection with write permissions. &#10;
    /// adModeShareExclusive 12 Prevents others from opening a connection.                       &#10;
    /// adModeShareDenyNone  16 Allows others to open a connection with any permissions.         &#10;
    /// adModeRecursive      0x400000 Used with adModeShareDenyNone, adModeShareDenyWrite, or adModeShareDenyRead to set permissions on all sub-records of the current Record.
    /// </field>
    /// <field name="Position" type="Number">Sets or returns the current position (in bytes) from the beginning of a Stream object</field>
    /// <field name="Size" type="String">Returns the size of an open Stream object</field>
    /// <field name="State" type="Number">Returns a value describing if the Stream object is open or closed &#10;
    /// adStateClosed      0 The object is closed                       &#10;
    /// adStateOpen        1 The object is open                         &#10;
    /// adStateConnecting  2 The object is connecting                   &#10;
    /// adStateExecuting   4 The object is executing a command          &#10;
    /// adStateFetching    8 The rows of the object are being retrieved &#10;
    /// </field>
    /// <field name="Type" type="Number">Sets or returns the type of data in a Stream object &#10;
    /// adTypeBinary 1 Binary data         &#10;
    /// adTypeText   2 Text data [Default] &#10;
    /// </field>
    this.CharSet = 'Unicode'
    this.LineSeparator = -1
    this.Mode = 0
    this.Position = 0
    this.Size = 0
    this.State = 0
    this.Type = 2
    this.EOS = false

    this.Cancel = function() {
      /// <summary>Cancels an execution of an Open call on a Stream object</summary>
    }

    this.Close = function() {
      /// <summary>Closes a Stream object</summary>
    }

    this.CopyTo = function(stream, chars) {
      /// <summary>Copies a specified number of characters/bytes from one Stream object into another Stream object</summary>
      /// <param name="stream">An object variable value that contains a reference to an open Stream object. The current Stream is copied to the destination Stream specified by DestStream. The destination Stream must already be open. If not, a run-time error occurs.</param>
      /// <param name="chars">Optional. An Integer value that specifies the number of bytes or characters to be copied from the current position in the source Stream to the destination Stream. The default value is –1, which specifies that all characters or bytes are copied from the current position to EOS.</param>
    }

    this.Flush = function() {
      /// <summary>Sends the contents of the Stream buffer to the associated underlying object</summary>
    }

    this.LoadFromFile = function(fileName) {
      /// <summary>Loads the contents of a file into a Stream object</summary>
      /// <param name="fileName">A String value that contains the name of a file to be loaded into the Stream. FileName can contain any valid path and name in UNC format. If the specified file does not exist, a run-time error occurs.</param>
    }

    this.Open = function(source, mode, options, username, password) {
      /// <summary>Opens a Stream object</summary>
      /// <param name="source">Optional. A Variant value that specifies the source of data for the Stream. Source may contain an absolute URL string that points to an existing node in a well-known tree structure, such as an e-mail or file system. A URL should be specified by using the URL keyword ("URL=scheme://server/folder"). Alternatively, Source may contain a reference to an already open Record object, which opens the default stream associated with the Record. If Source is not specified, a Stream is instantiated and opened, associated with no underlying source by default. For more information about URL schemes and their associated providers, see Absolute and Relative URLs.</param>
      /// <param name="mode">Optional. A ConnectModeEnum value that specifies the access mode for the resultant Stream (for example, read/write or read-only). Default value is adModeUnknown. See the Mode property for more information about access modes. If Mode is not specified, it is inherited by the source object. For example, if the source Record is opened in read-only mode, the Stream will also be opened in read-only mode by default.</param>
      /// <param name="options">Optional. A StreamOpenOptionsEnum value. Default value is adOpenStreamUnspecified.</param>
      /// <param name="username">Optional. A String value that contains the user identification that, if it is needed, accesses the Stream object.</param>
      /// <param name="password">Optional. A String value that contains the password that, if it is needed, accesses the Stream object.</param>
    }

    this.Read = function(bytes) {
      /// <summary>Reads the entire stream or a specified number of bytes from a binary Stream object</summary>
      /// <param name="bytes">Optional. A Long value that specifies the number of bytes to read from the file or the StreamReadEnum value adReadAll, which is the default.</param>
    }

    this.ReadText = function(chars) {
      /// <summary>Reads the entire stream, a line, or a specified number of characters from a text Stream object</summary>
      /// <param name="chars">Optional. A Long value that specifies the number of characters to read from the file, or a StreamReadEnum value. The default value is adReadAll.
      /// &#10; adReadAll  -1 Default. Reads all bytes from the stream, from the current position onwards to the EOS marker. This is the only valid StreamReadEnum value with binary streams (Type is adTypeBinary).
      /// &#10; adReadLine -2 Reads the next line from the stream (designated by the LineSeparator property).
      /// </param>
      /// <returns type="String"/>
    }

    this.SaveToFile = function(fileName, options) {
      /// <summary>Saves the binary contents of a Stream object to a file</summary>
      /// <param name="fileName">A String value that contains the fully-qualified name of the file to which the contents of the Stream will be saved. You can save to any valid local location, or any location you have access to via a UNC value.</param>
      /// <param name="options">A SaveOptionsEnum value that specifies whether a new file should be created by SaveToFile, if it does not already exist. Default value is adSaveCreateNotExists. With these options you can specify that an error occurs if the specified file does not exist. You can also specify that SaveToFile overwrites the current contents of an existing file.
      /// &#10; adSaveCreateNotExist  1 Default. Creates a new file if the file specified by the FileName parameter does not already exist.
      /// &#10; adSaveCreateOverWrite 2 Overwrites the file with the data from the currently open Stream object, if the file specified by the Filename parameter already exists. If the file specified by the Filename parameter does not exist, a new file is created.
      /// </param>
    }

    this.SetEOS = function() {
      /// <summary>Sets the current position to be the end of the stream (EOS)</summary>
    }

    this.SkipLine = function() {
      /// <summary>Skips a line when reading a text Stream</summary>
    }

    this.Write = function(buffer) {
      /// <summary>Writes binary data to a binary Stream object</summary>
      /// <param name="buffer">A Variant that contains an array of bytes to be written.</param>
    }

    this.WriteText = function(data, options) {
      /// <summary>Writes character data to a text Stream object</summary>
      /// <param name="data">A String value that contains the text in characters to be written.</param>
      /// <param name="options">Optional. A StreamWriteEnum value that specifies whether a line separator character must be written at the end of the specified string.
      /// &#10; adWriteChar 0 Default. Writes the specified text string (specified by the Data parameter) to the Stream object.
      /// &#10; adWriteLine 1 Writes a text string and a line separator character to a Stream object. If the LineSeparator property is not defined, then this returns a run-time error.
      /// </param>
    }

  }

}();

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
    /// <field name="xml"                type="String"         >Contains the XML representation of the node and all its descendants. Read-only</field>

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
    this.getAllResponseHeaders = function() {
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

+function() {

  ActiveXObject['scripting.filesystemobject'] = FileSystemObject

  function FileSystemObject() {
      /// <field name="Drives" type="Drives">Returns the Drives collection.</field>
    this.Drives = new Drives()
    this.BuildPath = function(path, name) {
      /// <summary>The BuildPath method inserts an additional path separator between the existing path and the name, only if necessary.</summary>
      /// <param name="path">Required. Existing path to which name is appended. Path can be absolute or relative and need not specify an existing folder.</param>
      /// <param name="name">Required. Name being appended to the existing path.</param>
      /// <returns type="String"/>
    }
    this.CopyFile = function(source, destination, overwrite) {
      /// <summary>Copies one or more files from one location to another.</summary>
      /// <param name="source">Required. Character string file specification, which can include wildcard characters, for one or more files to be copied.</param>
      /// <param name="destination">Required. Character string destination where the file or files from source are to be copied. Wildcard characters are not allowed. </param>
      /// <param name="overwrite">Optional. Boolean value that indicates if existing files are to be overwritten. If true, files are overwritten; if false, they are not. The default is true. Note that CopyFile will fail if destination has the read-only attribute set, regardless of the value of overwrite.</param>
    }
    this.CopyFolder = function(source, destination, overwrite) {
      /// <summary>Recursively copies a folder from one location to another.</summary>
      /// <param name="source">Required. Character string folder specification, which can include wildcard characters, for one or more folders to be copied. </param>
      /// <param name="destination">Required. Character string destination where the folder and subfolders from source are to be copied. Wildcard characters are not allowed.</param>
      /// <param name="overwrite">Optional. Boolean value that indicates if existing folders are to be overwritten. If true, files are overwritten; if false, they are not. The default is true.</param>
    }
    this.CreateFolder = function(foldername) {
      /// <summary>Creates a folder.</summary>
      /// <param name="foldername">Required. String expression that identifies the folder to create. </param>
    }
    this.CreateTextFile = function(filename, overwrite, unicode) {
      /// <summary>Creates a specified file name and returns a TextStream object that can be used to read from or write to the file.</summary>
      /// <param name="filename">Required. String expression that identifies the file to create.</param>
      /// <param name="overwrite">Optional. Boolean value that indicates whether you can overwrite an existing file. The value is true if the file can be overwritten, false if it can't be overwritten. If omitted, existing files are not overwritten. </param>
      /// <param name="unicode">Optional. Boolean value that indicates whether the file is created as a Unicode or ASCII file. The value is true if the file is created as a Unicode file, false if it's created as an ASCII file. If omitted, an ASCII file is assumed. </param>
      /// <returns type="TextStream"/>
      return new TextStream
    }
    this.DeleteFile = function(filespec, force) {
      /// <summary>Deletes a specified file.</summary>
      /// <param name="filespec">Required. The name of the file to delete. The filespec can contain wildcard characters in the last path component.</param>
      /// <param name="force">Optional. Boolean value that is true if files with the read-only attribute set are to be deleted; false (default) if they are not.</param>
    }
    this.DeleteFolder = function(folderspec, force) {
      /// <summary>Deletes a specified folder and its contents.</summary>
      /// <param name="folderspec">Required. The name of the folder to delete. The folderspec can contain wildcard characters in the last path component.</param>
      /// <param name="force">Optional. Boolean value that is true if folders with the read-only attribute set are to be deleted; false (default) if they are not.</param>
    }
    this.DriveExists = function(drivespec) {
      /// <summary>Returns True if the specified drive exists; False if it does not.&#10;For drives with removable media, the DriveExists method returns true even if there are no media present. Use the IsReady property of the Drive object to determine if a drive is ready.</summary>
      /// <param name="drivespec">Required. A drive letter or a complete path specification. </param>
      /// <returns type="Boolean"/>
    }
    this.FileExists = function(filespec) {
      /// <summary>Returns True if a specified file exists; False if it does not.</summary>
      /// <param name="filespec">Required. The name of the file whose existence is to be determined. A complete path specification (either absolute or relative) must be provided if the file isn't expected to exist in the current folder.</param>
      /// <returns type="Boolean"/>
    }
    this.FolderExists = function(folderspec) {
      /// <summary>Returns True if a specified folder exists; False if it does not.</summary>
      /// <param name="folderspec">Required. The name of the folder whose existence is to be determined. A complete path specification (either absolute or relative) must be provided if the folder isn't expected to exist in the current folder.</param>
      /// <returns type="Boolean"/>
    }
    this.GetAbsolutePathName = function(pathspec) {
      /// <summary>Returns a complete and unambiguous path from a provided path specification.</summary>
      /// <param name="pathspec">Required. Path specification to change to a complete and unambiguous path.&#10;
      /// "c:".........................."c:\mydocuments\reports"           &#10;
      /// "c:.."........................"c:\mydocuments"                   &#10;
      /// "c:\\"........................"c:\"                              &#10;
      /// "c:*.*\\may97"................"c:\mydocuments\reports\*.*\may97" &#10;
      /// "region1"....................."c:\mydocuments\reports\region1"   &#10;
      /// "c:\\..\\..\\mydocuments"....."c:\mydocuments" 
      /// </param>
      /// <returns type="String"/>
    }
    this.GetBaseName = function(path) {
      /// <summary>Returns a string containing the base name of the last component, less any file extension, in a path.</summary>
      /// <param name="path">Required. The path specification for the component whose base name is to be returned.</param>
      /// <returns type="String"/>
    }
    this.GetDrive = function(drivespec) {
      /// <summary>Returns a Drive object corresponding to the drive in a specified path.</summary>
      /// <param name="drivespec">Required. The drivespec argument can be a drive letter (c), a drive letter with a colon appended (c:), a drive letter with a colon and path separator appended (c:\), or any network share specification (\\computer2\share1).</param>
      /// <returns type="Drive"/>
      return new Drive
    }
    this.GetDriveName = function(path) {
      /// <summary>Returns a string containing the name of the drive for a specified path.</summary>
      /// <param name="path">Required. The path specification for the component whose drive name is to be returned</param>
      /// <returns type="String"/>
    }
    this.GetExtensionName = function(path) {
      /// <summary>Returns a string containing the extension name for the last component in a path.</summary>
      /// <param name="path">Required. The path specification for the component whose extension name is to be returned.</param>
      /// <returns type="String"/>
    }
    this.GetFile = function(filespec) {
      /// <summary>Returns a File object corresponding to the file in a specified path.</summary>
      /// <param name="filespec">Required. The filespec is the path (absolute or relative) to a specific file.</param>
      /// <returns type="File"/>
      return new File
    }
    this.GetFileName = function(pathspec) {
      /// <summary>Returns the last component of specified path that is not part of the drive specification.</summary>
      /// <param name="pathspec">Required. The path (absolute or relative) to a specific file.</param>
      /// <returns type="String"/>
    }
    this.GetFolder = function(folderspec) {
      /// <summary>Returns a Folder object corresponding to the folder in a specified path.</summary>
      /// <param name="folderspec">Required. The folderspec is the path (absolute or relative) to a specific folder.</param>
      /// <returns type="Folder"/>
      return new Folder
    }
    this.GetParentFolderName = function(path) {
      /// <summary>Returns a string containing the name of the parent folder of the last component in a specified path.</summary>
      /// <param name="path">Required. The path specification for the component whose parent folder name is to be returned.</param>
      /// <returns type="String"/>
    }
    this.GetSpecialFolder = function(folderspec) {
      /// <summary>Returns the special folder object specified.</summary>
      /// <param name="folderspec">Required. The name of the special folder to be returned. Can be any of the following.&#10;
      /// 0 WindowsFolder, The Windows folder contains files installed by the Windows operating system.                           &#10;
      /// 1 SystemFolder, The System folder contains libraries, fonts, and device drivers.                                        &#10;
      /// 2 TemporaryFolder, The Temp folder is used to store temporary files. Its path is found in the TMP environment variable. 
      /// </param>
      /// <returns type="Folder"/>
      return new Folder
    }
    this.GetTempName = function() {
      /// <summary>Returns a randomly generated temporary file or folder name that is useful for performing operations that require a temporary file or folder.</summary>
      /// <returns type="String"/>
    }
    this.MoveFile = function(source, destination) {
      /// <summary>Moves one or more files from one location to another.</summary>
      /// <param name="source">Required. The path to the file or files to be moved. The source argument string can contain wildcard characters in the last path component only. </param>
      /// <param name="destination">Required. The path where the file or files are to be moved. The destination argument can't contain wildcard characters. </param>
    }
    this.MoveFolder = function(source, destination) {
      /// <summary>Moves one or more folders from one location to another.</summary>
      /// <param name="source">Required. The path to the folder or folders to be moved. The source argument string can contain wildcard characters in the last path component only.</param>
      /// <param name="destination">Required. The path where the folder or folders are to be moved. The destination argument can't contain wildcard characters.</param>
    }
    this.OpenTextFile = function(filename, iomode, create, format) {
      /// <summary>Opens a specified file and returns a TextStream object that can be used to read from, write to, or append to the file.</summary>
      /// <param name="filename">Required. String expression that identifies the file to open.</param>
      /// <param name="iomode">Optional. Can be one of three constants: &#10;
      /// 1 ForReading, Open a file for reading only. You can't write to this file.                                          &#10;
      /// 2 ForWriting, Open a file for writing. If a file with the same name exists, its previous contents are overwritten. &#10;
      /// 8 ForAppending, Open a file and write to the end of the file. 
      /// </param>
      /// <param name="create">Optional. Boolean value that indicates whether a new file can be created if the specified filename doesn't exist. The value is True if a new file is created, False if it isn't created. If omitted, a new file isn't created.</param>
      /// <param name="format">Optional. One of three Tristate values used to indicate the format of the opened file. If omitted, the file is opened as ASCII.&#10;
      ///  0 TristateFalse, Open the file as ASCII.                      &#10;
      /// -1 TristateTrue, Open the file as Unicode.                     &#10;
      /// -2 TristateUseDefault, Open the file using the system default. 
      /// </param>
      /// <returns type="TextStream"/>
      return new TextStream
    }
  }

  function Drive() {
    /// <field name="AvailableSpace" type="Number">Returns the amount of space available to a user on the specified drive or network share.</field>
    /// <field name="DriveLetter" type="String">Returns the drive letter of a physical local drive or a network share. Read-only.</field>
    /// <field name="DriveType" type="Number">Returns a value indicating the type of a specified drive.&#10;0  Unknown&#10;1  Removable&#10;2  Fixed&#10;3  Network&#10;4  CD ROM&#10;5  RAM Disk</field>
    /// <field name="FileSystem" type="String">Returns the type of file system in use for the specified drive. FAT, NTFS, CDFS</field>
    /// <field name="FreeSpace" type="Number">Returns the amount of free space available to a user on the specified drive or network share. Read-only.</field>
    /// <field name="IsReady" type="Boolean">Returns True if the specified drive is ready; False if it is not.</field>
    /// <field name="Path" type="String">Returns the path for a specified file, folder, or drive.</field>
    /// <field name="RootFolder" type="Folder">Returns a Folder object representing the root folder of a specified drive. Read-only</field>
    /// <field name="SerialNumber" type="Number">Returns the decimal serial number used to uniquely identify a disk volume.</field>
    /// <field name="ShareName" type="String">Returns the network share name for a specified drive.</field>
    /// <field name="TotalSize" type="Number">Returns the total space, in bytes, of a drive or network share.</field>
    /// <field name="VolumeName" type="String">Sets or returns the volume name of the specified drive. Read/write.</field>
    this.AvailableSpace = 0
    this.DriveLetter    = ''
    this.DriveType      = 0
    this.FileSystem     = ''
    this.FreeSpace      = 0
    this.IsReady        = false
    this.Path           = ''
    this.RootFolder     = new Folder
    this.SerialNumber   = 0.0
    this.ShareName      = ''
    this.TotalSize      = 0
    this.VolumeName     = ''
  }

  function File(parentFolder) {
    /// <field name="Attributes" type="Number">Sets or returns the attributes of files or folders. Read/write or read-only, depending on the attribute.&#10;
    /// 0 Normal file. No attributes are set.                                    &#10;
    /// 1 Read-only file. Attribute is read/write.                               &#10;
    /// 2 Hidden file. Attribute is read/write.                                  &#10;
    /// 4 System file. Attribute is read/write.                                  &#10;
    /// 8 Disk drive volume label. Attribute is read-only.                       &#10;
    /// 16 Directory or Folder. Attribute is read-only.                          &#10;
    /// 32 Archive, File has changed since last backup. Attribute is read/write. &#10;
    /// 64 Alias, Link or shortcut. Attribute is read-only.                      &#10;
    /// 128 Compressed file. Attribute is read-only.
    /// </field>
    /// <field name="DateCreated" type="String">Read-only.</field>
    /// <field name="DateLastAccessed" type="String">Read-only.</field>
    /// <field name="DateLastModified" type="String">Read-only.</field>
    /// <field name="Drive" type="String">Returns the drive letter of the drive on which the specified file or folder resides. Read-only.</field>
    /// <field name="Name" type="String">Sets or returns the name of a specified file or folder. Read/write.</field>
    /// <field name="ParentFolder" type="">Returns the folder object for the parent of the specified file or folder. Read-only.</field>
    /// <field name="Path" type="String">Returns the path for a specified file, folder, or drive.</field>
    /// <field name="ShortName" type="String">Returns the short name used by programs that require the earlier 8.3 naming convention.</field>
    /// <field name="ShortPath" type="String">Returns the short path used by programs that require the earlier 8.3 file naming convention.</field>
    /// <field name="Size" type="">For files, returns the size, in bytes, of the specified file. For folders, returns the size, in bytes, of all files and subfolders contained in the folder.</field>
    /// <field name="Type" type="String">Returns information about the type of a file or folder. For example, for files ending in .TXT, "Text Document" is returned.</field>
    this.Attributes       = 0
    this.DateCreated      = ''
    this.DateLastAccessed = ''
    this.DateLastModified = ''
    this.Drive            = ''
    this.Name             = ''
    this.ParentFolder     = parentFolder || new Folder(this)
    this.Path             = ''
    this.ShortName        = ''
    this.ShortPath        = ''
    this.Size             = 0
    this.Type             = ''
    this.Copy = function(destination, overwrite) {
      /// <summary>Copies a specified file or folder from one location to another.</summary>
      /// <param name="destination">Destination where the file or folder is to be copied. Wildcard characters are not allowed.</param>
      /// <param name="overwrite">Optional. Boolean value that is True (default) if existing file is to be overwritten; False if they are not.</param>
    }
    this.Delete = function(force) {
      /// <summary>Deletes a specified file.</summary>
      /// <param name="force">Optional. Boolean value that is True if file with the read-only attribute set is to be deleted; False (default) if it is not.</param>
    }
    this.Move = function(destination) {
      /// <summary>Moves a specified file from one location to another.</summary>
      /// <param name="destination">Required. Destination where the file or folder is to be moved. Wildcard characters are not allowed.</param>
    }
    this.OpenAsTextStream = function(iomode, format) {
      /// <summary>Opens a specified file and returns a TextStream object that can be used to read from, write to, or append to the file.</summary>
      /// <param name="iomode">Optional. Indicates input/output mode. Can be one of three constants:                          &#10;
      /// 1 ForReading, Open a file for reading only. You can't write to this file.                                           &#10;
      /// 2 ForWriting, Open a file for writing. If a file with the same name exists, its previous contents are overwritten.  &#10;
      /// 8 ForAppending, Open a file and write to the end of the file. 
      /// </param>
      /// <param name="format">Optional. One of three Tristate values used to indicate the format of the opened file. If omitted, the file is opened as ASCII.&#10;
      ///  0 TristateFalse, Open the file as ASCII.                      &#10;
      /// -1 TristateTrue, Open the file as Unicode.                     &#10;
      /// -2 TristateUseDefault, Open the file using the system default. 
      /// </param>
      /// <returns type="TextStream"/>
      return new TextStream
    }
  }

  function Folder(parentFolder) {
    /// <field name="Attributes" type="Number">Sets or returns the attributes of files or folders. Read/write or read-only, depending on the attribute.&#10;
    /// 0 Normal file. No attributes are set.                                    &#10;
    /// 1 Read-only file. Attribute is read/write.                               &#10;
    /// 2 Hidden file. Attribute is read/write.                                  &#10;
    /// 4 System file. Attribute is read/write.                                  &#10;
    /// 8 Disk drive volume label. Attribute is read-only.                       &#10;
    /// 16 Directory or Folder. Attribute is read-only.                          &#10;
    /// 32 Archive, File has changed since last backup. Attribute is read/write. &#10;
    /// 64 Alias, Link or shortcut. Attribute is read-only.                      &#10;
    /// 128 Compressed file. Attribute is read-only.
    /// </field>
    /// <field name="DateCreated" type="String">Read-only.</field>
    /// <field name="DateLastAccessed" type="String">Read-only.</field>
    /// <field name="DateLastModified" type="String">Read-only.</field>
    /// <field name="Drive" type="String">Returns the drive letter of the drive on which the specified file or folder resides. Read-only.</field>
    /// <field name="Files" type="Files">Returns a Files collection consisting of all File objects contained in the specified folder, including those with hidden and system file attributes set.</field>
    /// <field name="IsRootFolder" type="Boolean">Returns True if the specified folder is the root folder; False if it is not.</field>
    /// <field name="Name" type="String">Sets or returns the name of a specified file or folder. Read/write.</field>
    /// <field name="ParentFolder" type="Folder">Returns the folder object for the parent of the specified file or folder. Read-only.</field>
    /// <field name="Path" type="String">Returns the path for a specified file, folder, or drive.</field>
    /// <field name="ShortName" type="String">Returns the short name used by programs that require the earlier 8.3 naming convention.</field>
    /// <field name="ShortPath" type="String">Returns the short path used by programs that require the earlier 8.3 file naming convention.</field>
    /// <field name="Size" type="Number">For files, returns the size, in bytes, of the specified file. For folders, returns the size, in bytes, of all files and subfolders contained in the folder.</field>
    /// <field name="SubFolders" type="Folders">Returns a Folders collection consisting of all folders contained in a specified folder, including those with hidden and system file attributes set.</field>
    /// <field name="Type" type="String">Returns information about the type of a file or folder. For example, for files ending in .TXT, "Text Document" is returned.</field>
    this.Attributes       = 0
    this.DateCreated      = ''
    this.DateLastAccessed = ''
    this.DateLastModified = ''
    this.Drive            = ''
    this.Files            = new Files
    this.IsRootFolder     = false
    this.Name             = ''
    this.ParentFolder     = parentFolder || new Folder(this)
    this.Path             = ''
    this.ShortName        = ''
    this.ShortPath        = ''
    this.Size             = 0
    this.SubFolders       = new Folders
    this.Type             = ''
    this.Copy = function(destination, overwrite) {
      /// <summary>Copies a specified file or folder from one location to another.</summary>
      /// <param name="destination">Required. Destination where the file or folder is to be copied. Wildcard characters are not allowed. </param>
      /// <param name="overwrite">Optional. Boolean value that is True (default) if existing files or folders are to be overwritten; False if they are not. </param>
    }
    this.Delete = function(force) {
      /// <summary>Deletes a specified file or folder.</summary>
      /// <param name="force">Optional. Boolean value that is True if files or folders with the read-only attribute set are to be deleted; False (default) if they are not. </param>
    }
    this.Move = function(destination) {
      /// <summary>Moves a specified file or folder from one location to another.</summary>
      /// <param name="destination">Required. Destination where the file or folder is to be moved. Wildcard characters are not allowed. </param>
    }
  }

  function TextStream() {
    /// <field name="AtEndOfLine" type="Boolean">Returns true if the file pointer is positioned immediately before the end-of-line marker in a TextStream file; false if it is not. Read-only.</field>
    /// <field name="AtEndOfStream" type="Boolean">Returns true if the file pointer is at the end of a TextStream file; false if it is not. Read-only.</field>
    /// <field name="Column" type="Number">Read-only property that returns the column number of the current character position in a TextStream file.</field>
    /// <field name="Line" type="Number">Read-only property that returns the current line number in a TextStream file.</field>
    this.AtEndOfLine   = true
    this.AtEndOfStream = true
    this.Column        = 0
    this.Line          = 0
    this.Close = function() {
      /// <summary>Closes an open TextStream file.</summary>
    }
    this.Read = function(characters) {
      /// <summary>Reads a specified number of characters from a TextStream file and returns the resulting string.</summary>
      /// <param name="characters">Required. Number of characters you want to read from the file.</param>
      /// <returns type="String"/>
    }
    this.ReadAll = function() {
      /// <summary>Reads an entire TextStream file and returns the resulting string.</summary>
      /// <returns type="String"/>
    }
    this.ReadLine = function() {
      /// <summary>Reads an entire line (up to, but not including, the newline character) from a TextStream file and returns the resulting string.</summary>
      /// <returns type="String"/>
    }
    this.Skip = function(characters) {
      /// <summary>Skips a specified number of characters when reading a TextStream file.</summary>
      /// <param name="characters">Required. Number of characters to skip when reading a file.</param>
    }
    this.SkipLine = function() {
      /// <summary>Skips the next line when reading a TextStream file.</summary>
    }
    this.Write = function(string) {
      /// <summary>Writes a specified string to a TextStream file.</summary>
      /// <param name="string">Required. The text you want to write to the file.</param>
    }
    this.WriteBlankLines = function(lines) {
      /// <summary>Writes a specified number of newline characters to a TextStream file.</summary>
      /// <param name="lines">Required. Number of newline characters you want to write to the file.</param>
    }
    this.WriteLine = function(string) {
      /// <summary>Writes a specified string and newline character to a TextStream file.</summary>
      /// <param name="string">Optional. The text you want to write to the file. If omitted, a newline character is written to the file.</param>
    }
  }

  function Drives() {
    this.Count = 0
    this.Item = function(key) {
      /// <summary>Returns an item based on the specified key.</summary>
      /// <returns type="Drive"/>
      return new Drive
    }
  }

  function Files() {
    this.Count = 0
    this.Item = function(key) {
      /// <summary>Returns an item based on the specified key.</summary>
      /// <returns type="File"/>
      return new File
    }
  }

  function Folders() {
    this.Count = 0
    this.Add = function(foldername) {
      /// <summary>Adds a new folder to a Folders collection.</summary>
      /// <param name="foldername">Required. The name of the new Folder being added.</param>
    }
    this.Item = function(key) {
      /// <summary>Returns an item based on the specified key.</summary>
      /// <returns type="Folder"/>
      return new Folder
    }
  }

}();

+function() {

  ActiveXObject['shell.application'] = ShellApplication

  function ShellApplication() {
    this.AddToRecent = function(file, category) {
      /// <summary>Adds a file to the most recently used (MRU) list.</summary>
      /// <param name="file">string, path of the file to add to the list of recently used documents&#10;null, to clear the recent documents folder</param>
      /// <param name="category">string, name of the category in which ot place the file (optional)</param>
    }
    this.BrowseForFolder = function(hwnd, title, options, rootFolder) {
      /// <summary>Creates a dialog box that enables the user to select a folder and then returns the selected folder's Folder object.</summary>
      /// <param name="hwnd">number, handle of the parent window of the dialog box, or zero.</param>
      /// <param name="title">string, represents the title displayed inside the Browse dialog box.</param>
      /// <param name="options">number, zero or a combination of the ulFlags in BROWSEINFO</param>
      /// <param name="rootFolder">user can not browse higher than this folder &#10;String, specifying a path &#10;Number, one of the ShellSpecialFolderConstants (ssf) &#10;Undefined, will use DESKTOP</param>
      /// <returns type="Folder"/>
      return new Folder
    }
    this.CanStartStopService = function() {
      /// <summary>Determines if the current user can start and stop the named service.</summary>
      /// <param name=""></param>
    }
    this.CascadeWindows = function() {
      /// <summary>Cascades all of the windows on the desktop. This method has the same effect as right-clicking the taskbar and selecting Cascade Windows.</summary>
      /// <param name=""></param>
    }
    this.ControlPanelItem = function() {
      /// <summary>Runs the specified Control Panel (*.cpl) application. If the application is already open, it will activate the running instance.</summary>
      /// <param name=""></param>
    }
    this.EjectPC = function() {
      /// <summary>Ejects the computer from its docking station. This is the same as clicking the Start menu and selecting Eject PC, if your computer supports this command.</summary>
      /// <param name=""></param>
    }
    this.Explore = function() {
      ///<summary>Opens a specified folder in a Windows Explorer window.</summary>
      ///<param name=""></param>
    }
    this.ExplorerPolicy = function() {
      /// <summary>Gets the value for a specified Internet Explorer policy.</summary>
      /// <param name=""></param>
    }
    this.FileRun = function() {
      /// <summary>Displays the Run dialog to the user. This method has the same effect as clicking the Start menu and selecting Run.</summary>
      /// <param name=""></param>
    }
    this.FindComputer = function() {
      /// <summary>Displays the Search Results: Computers dialog box. The dialog box shows the result of the search for a specified computer.</summary>
      /// <param name=""></param>
    }
    this.FindFiles = function() {
      /// <summary>Displays the Find: All Files dialog box. This is the same as clicking the Start menu and then selecting Search (or its equivalent under systems earlier than Windows XP.</summary>
      /// <param name=""></param>
    }
    this.FindPrinter = function() {
      /// <summary>Displays the Find Printer dialog box.</summary>
      /// <param name=""></param>
    }
    this.GetSetting = function() {
      /// <summary>Retrieves a global Shell setting.</summary>
      /// <param name=""></param>
    }
    this.GetSystemInformation = function() {
      /// <summary>Retrieves system information.</summary>
      /// <param name=""></param>
    }
    this.Help = function() {
      /// <summary>Displays the Windows Help and Support Center. This method has the same effect as clicking the Start menu and selecting Help and Support.</summary>
      /// <param name=""></param>
    }
    this.IsRestricted = function() {
      /// <summary>Retrieves a group's restriction setting from the registry.</summary>
      /// <param name=""></param>
    }
    this.IsServiceRunning = function() {
      /// <summary>Returns a value that indicates whether a particular service is running.</summary>
      /// <param name=""></param>
    }
    this.MinimizeAll = function() {
      /// <summary>Minimizes all of the windows on the desktop. This method has the same effect as right-clicking the taskbar and selecting Minimize All Windows on older systems or clicking the Show Desktop icon in the Quick Launch area of the taskbar in Windows 2000 or Windows XP.</summary>
      /// <param name=""></param>
    }
    this.NameSpace = function(dir) {
      /// <summary>Creates and returns a Folder object for the specified folder.</summary>
      /// <param name="dir">String, path of the folder &#10;Number, a value from the ShellSpecialFolderConstants (ssf)</param>
      return new Folder
    }
    this.Open = function() {
      /// <summary>Opens the specified folder.</summary>
      /// <param name=""></param>
    }
    this.RefreshMenu = function() {
      /// <summary>Refreshes the contents of the Start menu. Used only with systems preceding Windows XP.</summary>
      /// <param name=""></param>
    }
    this.ServiceStart = function() {
      /// <summary>Starts a named service.</summary>
      /// <param name=""></param>
    }
    this.ServiceStop = function() {
      /// <summary>Stops a named service.</summary>
      /// <param name=""></param>
    }
    this.SetTime = function() {
      /// <summary>Displays the Date and Time Properties dialog box. This method has the same effect as right-clicking the clock in the taskbar status area and selecting Adjust Date/Time.</summary>
      /// <param name=""></param>
    }
    this.ShellExecute = function() {
      /// <summary>Performs a specified operation on a specified file.</summary>
      /// <param name=""></param>
    }
    this.ShowBrowserBar = function() {
      /// <summary>Displays a browser bar.</summary>
      /// <param name=""></param>
    }
    this.ShutdownWindows = function() {
      /// <summary>Displays the Shut Down Windows dialog box. This is the same as clicking the Start menu and selecting Shut Down.</summary>
      /// <param name=""></param>
    }
    this.Suspend = function() {
      /// <summary></summary>
      /// <param name=""></param>
    }
    this.TileHorizontally = function() {
      /// <summary>Tiles all of the windows on the desktop horizontally. This method has the same effect as right-clicking the taskbar and selecting Tile Windows Horizontally.</summary>
      /// <param name=""></param>
    }
    this.TileVertically = function() {
      /// <summary>Tiles all of the windows on the desktop vertically. This method has the same effect as right-clicking the taskbar and selecting Tile Windows Vertically.</summary>
      /// <param name=""></param>
    }
    this.TrayProperties = function() {
      /// <summary>Displays the Taskbar and Start Menu Properties dialog box. This method has the same effect as right-clicking the taskbar and selecting Properties.</summary>
      /// <param name=""></param>
    }
    this.UndoMinimizeALL = function() {
      /// <summary>Restores all desktop windows to the same state they were in before the last MinimizeAll command. This method has the same effect as right-clicking the taskbar and selecting Undo Minimize All Windows on older systems or a second clicking of the Show Desktop icon in the Quick Launch area of the taskbar in Windows 2000 or Windows XP.</summary>
      /// <param name=""></param>
    }
    this.Windows = function() {
      /// <summary>Creates and returns a ShellWindows object. This object represents a collection of all of the open windows that belong to the Shell.</summary>
      /// <param name=""></param>
    }
    this.WindowsSecurity = function() {
      /// <summary>Displays the Windows Security dialog box.</summary>
      /// <param name=""></param>
    }
    this.WindowSwitcher = function() {
      /// <summary>Displays your open windows in a 3D stack that you can flip through.</summary>
      /// <param name=""></param>
    }

    this.ulFlags = {
      // Only return file system directories. If the user selects folders that are not part of the file system, the OK button is grayed. Note  The OK button remains enabled for "\\server" items, as well as "\\server\share" and directory items. However, if the user selects a "\\server" item, passing the PIDL returned by SHBrowseForFolder to SHGetPathFromIDList fails.
      BIF_RETURNONLYFSDIRS    :     1,
      // Do not include network folders below the domain level in the dialog box's tree view control.
      BIF_DONTGOBELOWDOMAIN   :     2,
      // Include a status area in the dialog box. The callback function can set the status text by sending messages to the dialog box. This flag is not supported when BIF_NEWDIALOGSTYLE is specified.
      BIF_STATUSTEXT          :     4,
      // Only return file system ancestors. An ancestor is a subfolder that is beneath the root folder in the namespace hierarchy. If the user selects an ancestor of the root folder that is not part of the file system, the OK button is grayed.
      BIF_RETURNFSANCESTORS   :     8,
      // Version 4.71. Include an edit control in the browse dialog box that allows the user to type the name of an item.
      BIF_EDITBOX             :    10,
      // Version 4.71. If the user types an invalid name into the edit box, the browse dialog box calls the application's BrowseCallbackProc with the BFFM_VALIDATEFAILED message. This flag is ignored if BIF_EDITBOX is not specified.
      BIF_VALIDATE            :    20,
      // Version 5.0. Use the new user interface. Setting this flag provides the user with a larger dialog box that can be resized. The dialog box has several new capabilities, including: drag-and-drop capability within the dialog box, reordering, shortcut menus, new folders, delete, and other shortcut menu commands. Note  If COM is initialized through CoInitializeEx with the COINIT_MULTITHREADED flag set, SHBrowseForFolder fails if BIF_NEWDIALOGSTYLE is passed.
      BIF_NEWDIALOGSTYLE      :    40,
      // Version 5.0. The browse dialog box can display URLs. The BIF_USENEWUI and BIF_BROWSEINCLUDEFILES flags must also be set. If any of these three flags are not set, the browser dialog box rejects URLs. Even when these flags are set, the browse dialog box displays URLs only if the folder that contains the selected item supports URLs. When the folder's IShellFolder::GetAttributesOf method is called to request the selected item's attributes, the folder must set the SFGAO_FOLDER attribute flag. Otherwise, the browse dialog box will not display the URL.
      BIF_BROWSEINCLUDEURLS   :    80,
      // Version 5.0. Use the new user interface, including an edit box. This flag is equivalent to BIF_EDITBOX | BIF_NEWDIALOGSTYLE. Note  If COM is initialized through CoInitializeEx with the COINIT_MULTITHREADED flag set, SHBrowseForFolder fails if BIF_USENEWUI is passed.
      BIF_USENEWUI            :    50,
      // Version 6.0. When combined with BIF_NEWDIALOGSTYLE, adds a usage hint to the dialog box, in place of the edit box. BIF_EDITBOX overrides this flag.
      BIF_UAHINT              :   100,
      // Version 6.0. Do not include the New Folder button in the browse dialog box.
      BIF_NONEWFOLDERBUTTON   :   200,
      // Version 6.0. When the selected item is a shortcut, return the PIDL of the shortcut itself rather than its target.
      BIF_NOTRANSLATETARGETS  :   400,
      // Only return computers. If the user selects anything other than a computer, the OK button is grayed.
      BIF_BROWSEFORCOMPUTER   :  1000,
      // Only allow the selection of printers. If the user selects anything other than a printer, the OK button is grayed. In Windows XP and later systems, the best practice is to use a Windows XP-style dialog, setting the root of the dialog to the Printers and Faxes folder (CSIDL_PRINTERS).
      BIF_BROWSEFORPRINTER    :  2000,
      // Version 4.71. The browse dialog box displays files as well as folders.
      BIF_BROWSEINCLUDEFILES  :  4000,
      // Version 5.0. The browse dialog box can display sharable resources on remote systems. This is intended for applications that want to expose remote shares on a local system. The BIF_NEWDIALOGSTYLE flag must also be set.
      BIF_SHAREABLE           :  8000,
      // Windows 7 and later. Allow folder junctions such as a library or a compressed file with a .zip file name extension to be browsed.
      BIF_BROWSEFILEJUNCTIONS : 10000
    }

    this.ssf = {
      // File system directory that corresponds to the user's non-localized Startup program group.
      ALTSTARTUP         : 29,
      // File system directory that serves as a common repository for application-specific data. A typical path is C:\Documents and Settings\username\Application Data.
      APPDATA            : 26,
      // Virtual folder that contains the objects in the user's Recycle Bin.
      BITBUCKET          : 10,
      // File system directory that corresponds to the non-localized Startup program group for all users. Valid only for Windows NT systems.
      COMMONALTSTARTUP   : 30,
      // Application data for all users. A typical path is C:\Documents and Settings\All Users\Application Data.
      COMMONAPPDATA      : 35,
      // File system directory that contains files and folders that appear on the desktop for all users. A typical path is C:\Documents and Settings\All Users\Desktop. Valid only for Windows NT systems.
      COMMONDESKTOPDIR   : 25,
      // File system directory that serves as a common repository for the favorite URLs shared by all users. Valid only for Windows NT systems.
      COMMONFAVORITES    : 31,
      // File system directory that contains the directories for the common program groups that appear on the Start menu for all users. A typical path is C:\Documents and Settings\All Users\Start Menu\Programs. Valid only for Windows NT systems.
      COMMONPROGRAMS     : 23,
      // File system directory that contains the programs and folders that appear on the Start menu for all users. A typical path is C:\Documents and Settings\All Users\Start Menu. Valid only for Windows NT systems.
      COMMONSTARTMENU    : 22,
      // File system directory that contains the programs that appear in the Startup folder for all users. A typical path is C:\Documents and Settings\All Users\Microsoft\Windows\Start Menu\Programs\StartUp. Valid only for Windows NT systems.
      COMMONSTARTUP      : 24,
      // Virtual folder that contains icons for the Control Panel applications.
      CONTROLS           :  3,
      // File system directory that serves as a common repository for Internet cookies. A typical path is C:\Documents and Settings\username\Application Data\Microsoft\Windows\Cookies.
      COOKIES            : 33,
      // Windows desktop—the virtual folder that is the root of the namespace.
      DESKTOP            :  0,
      // File system directory used to physically store the file objects that are displayed on the desktop. It is not to be confused with the desktop folder itself, which is a virtual folder. A typical path is C:\Documents and Settings\username\Desktop.
      DESKTOPDIRECTORY   : 16,
      // My Computer—the virtual folder that contains everything on the local computer: storage devices, printers, and Control Panel. This folder can also contain mapped network drives.
      DRIVES             : 17,
      // File system directory that serves as a common repository for the user's favorite URLs. A typical path is C:\Documents and Settings\username\Favorites.
      FAVORITES          :  6,
      // Virtual folder that contains installed fonts. A typical path is C:\Windows\Fonts.
      FONTS              : 20,
      // File system directory that serves as a common repository for Internet history items.
      HISTORY            : 34,
      // File system directory that serves as a common repository for temporary Internet files. A typical path is C:\Users\username\AppData\Local\Microsoft\Windows\Temporary Internet Files.
      INTERNETCACHE      : 32,
      // File system directory that serves as a data repository for local (non-roaming) applications. A typical path is C:\Users\username\AppData\Local.
      LOCALAPPDATA       : 28,
      // My Pictures folder. A typical path is C:\Users\username\Pictures.
      MYPICTURES         : 39,
      // A file system folder that contains any link objects in the My Network Places virtual folder. It is not the same as ssfNETWORK, which represents the network namespace root. A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Network Shortcuts.
      NETHOOD            : 19,
      // Network Neighborhood—the virtual folder that represents the root of the network namespace hierarchy.
      NETWORK            : 18,
      // File system directory that serves as a common repository for a user's documents. A typical path is C:\Users\username\Documents.
      PERSONAL           :  5,
      // Virtual folder that contains installed printers.
      PRINTERS           :  4,
      // File system directory that contains any link objects in the Printers virtual folder. A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Printer Shortcuts.
      PRINTHOOD          : 27,
      // User's profile folder.
      PROFILE            : 40,
      // Program Files folder. A typical path is C:\Program Files.
      PROGRAMFILES       : 38,
      // Program Files folder. A typical path is C:\Program Files, or C:\Program Files (X86) on a 64-bit computer.
      PROGRAMFILESx86    : 48,
      // File system directory that contains the user's program groups (which are also file system directories). A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs.
      PROGRAMS           :  2,
      // File system directory that contains the user's most recently used documents. A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Recent.
      RECENT             :  8,
      // File system directory that contains Send To menu items. A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\SendTo.
      SENDTO             :  9,
      // File system directory that contains Start menu items. A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Start Menu.
      STARTMENU          : 11,
      // File system directory that corresponds to the user's Startup program group. The system starts these programs whenever any user first logs into their profile after a reboot. A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\StartUp.
      STARTUP            :  7,
      // The System folder. A typical path is C:\Windows\System32.
      SYSTEM             : 37,
      // System folder. A typical path is C:\Windows\System32, or C:\Windows\Syswow32 on a 64-bit computer.
      SYSTEMx86          : 41,
      // File system directory that serves as a common repository for document templates.
      TEMPLATES          : 21,
      // Windows directory. This corresponds to the %windir% or %SystemRoot% environment variables. A typical path is C:\Windows.
      WINDOWS            : 36
    }
  }

  function Folder() {
    // http://msdn.microsoft.com/en-us/library/bb787868%28v=vs.85%29.aspx
    // <field name="ParentFolder" type="Folder">Contains the parent Folder object.</field>
    // <field name="Title" type="String">Contains the title of the folder.</field>
    this.ParentFolder = new Folder
    this.Title = ''
    this.CopyHere = function() {
      /// <summary>Copies an item or items to a folder.</summary>
      /// <param name=""></param>
    }
    this.GetDetailsOf = function() {
      /// <summary>Retrieves details about an item in a folder. For example, its size, type, or the time of its last modification.</summary>
      /// <param name=""></param>
    }
    this.Items = function() {
      /// <summary>Retrieves a FolderItems object that represents the collection of items in the folder.</summary>
      /// <returns type="FolderItems"/>
      return new FolderItems
    }
    this.MoveHere = function() {
      /// <summary>Moves an item or items to this folder.</summary>
      /// <param name=""></param>
    }
    this.NewFolder = function() {
      /// <summary>Creates a new folder.</summary>
      /// <param name=""></param>
    }
    this.ParseName = function() {
      /// <summary>Creates and returns a FolderItem object that represents a specified item.</summary>
      /// <param name=""></param>
    }
  }

  function FolderItem() {
    // http://msdn.microsoft.com/en-us/library/bb787810%28v=vs.85%29.aspx
    /// <field name="GetFolder" type="Folder">Contains the item's Folder object, if the item is a folder.</field>
    /// <field name="GetLink" type="Link">Contains the item's ShellLinkObject object, if the item is a shortcut.</field>
    /// <field name="IsBrowsable" type="Boolean">Indicates if the item can be hosted inside a browser or Windows Explorer frame.</field>
    /// <field name="IsFileSystem" type="Boolean">Indicates if the item is part of the file system.</field>
    /// <field name="IsFolder" type="Boolean">Indicates if the item is a folder.</field>
    /// <field name="IsLink" type="Boolean">Indicates whether the item is a shortcut.</field>
    /// <field name="ModifyDate" type="String">Sets or gets the date and time that a file was last modified. ModifyDate can be used to retrieve the date and time that a folder was last modified, but cannot set it. "01/01/1900 6:05:00 PM"</field>
    /// <field name="Name" type="String">Sets or gets the item's name.</field>
    /// <field name="Parent" type="Folder">Gets an object that represents the parent of the item.</field>
    /// <field name="Path" type="String">Contains the item's full path and name.</field>
    /// <field name="Size" type="Number">Contains the item's size.</field>
    /// <field name="Type" type="Number">Contains a string representation of the item's type.</field>
    this.GetFolder = new Folder
    this.GetLink = new Link
    this.IsBrowsable = false
    this.IsFileSystem = false
    this.IsFolder = false
    this.IsLink = false
    this.ModifyDate = ''
    this.Name = ''
    this.Parent = null
    this.Path = ''
    this.Size = 0
    this.Type = 0
    this.InvokeVerb = function(verb) {
      /// <summary>Executes a verb on the item.</summary>
      /// <param name="verb">
      /// A string that specifies the verb to be executed.
      /// It must be one of the values returned by the item's FolderItemVerb.Name property.
      /// If no verb is specified, the default verb will be invoked.
      /// Typically "open"
      /// </param>
    }
    this.Verbs = function() {
      /// <summary>Retrieves the item's FolderItemVerbs object. This object is the collection of verbs that can be executed on the item.</summary>
      /// <returns type="FolderItemVerbs"/>
      return new FolderItemVerbs
    }
  }

  function FolderItems() {
    // http://msdn.microsoft.com/en-us/library/bb787800%28v=vs.85%29.aspx
    this.Count = 0
    this.Item = function(index) {
      /// <summary>Retrieves the FolderItem object for a specified item in the collection.</summary>
      /// <param name="index">number, zero-based index of the item to retrieve. This value must be less than the value of the Count property.</param>
      /// <returns type="FolderItem"/>
      return new FolderItem
    }
  }

  function FolderItemVerb() {
    // http://msdn.microsoft.com/en-us/library/bb774172%28v=vs.85%29.aspx
    /// <field name="Name" type="String">Contains the verb's name.</field>
    this.Name = ''
    this.DoIt = function() {
      /// <summary>Executes a verb on the FolderItem associated with the verb.</summary>
    }
  }

  function FolderItemVerbs() {
    // http://msdn.microsoft.com/en-us/library/bb774160%28v=vs.85%29.aspx
    this.Count = 0
    this.Item = function(index) {
      /// <summary>Retrieves the FolderItem object for a specified item in the collection.</summary>
      /// <param name="index">number, zero-based index of the item to retrieve. This value must be less than the value of the Count property.</param>
      /// <returns type="FolderItemVerb"/>
      return new FolderItemVerb
    }
  }

  function Link() {
    // http://msdn.microsoft.com/en-us/library/bb774004%28v=vs.85%29
    // <field name="Arguments" type="String">Contains a link's arguments.</field>
    // <field name="Description" type="String">Gets or sets the description of the link.</field>
    // <field name="Hotkey" type="String">Gets or sets the keyboard shortcut for the link.</field>
    // <field name="Path" type="String">Gets or sets the path to the link object.</field>
    // <field name="ShowCommand" type="String">Gets or sets the initial display state (sized, minimized, or maximized) of the link's command.</field>
    // <field name="WorkingDirectory" type="String">Gets or sets the working directory specified in the link.</field>
    this.Arguments = ''
    this.Description = ''
    this.Hotkey = ''
    this.Path = ''
    this.ShowCommand = ''
    this.WorkingDirectory = ''
    this.GetIconLocation = function() {
      /// <summary>Gets the location of the icon assigned to the link.</summary>
    }
    this.Resolve = function() {
      /// <summary>Looks for the target of a Shell link, even if the target has been moved or renamed.</summary>
    }
    this.Save = function() {
      /// <summary>Saves all changes to the link.</summary>
    }
    this.SetIconLocation = function() {
      /// <summary>Sets the location of the icon assigned to the link.</summary>
    }
  }

}();

+function() {

  ActiveXObject['wscript.network'] = WScriptNetwork

  function WScriptNetwork() {
    /// <field name="ComputerName" type="String">Returns the name of the computer system.</field>
    /// <field name="UserDomain" type="String">Returns a user's domain name.</field>
    /// <field name="UserName" type="String">Returns the name of a user.</field>
    this.ComputerName = ''
    this.UserDomain = ''
    this.UserName = ''
    this.AddWindowsPrinterConnection = function(path) {
      /// <summary>Adds a Windows-based printer connection to your computer system.&#10;Using this method is similar to using the Printer option on Control Panel to add a printer connection. Unlike the AddPrinterConnection method, this method allows you to create a printer connection without directing it to a specific port, such as LPT1. If the connection fails, an error is thrown.</summary>
      /// <param name="path">string, the path to the printer connection.&#10;"\\\\printserv\\DefaultPrinter"</param>
    }
    this.AddPrinterConnection = function(local, remote, updateProfile, user, password) {
      /// <summary>Adds a remote MS-DOS-based printer connection to your computer system.</summary>
      /// <param name="local">String value indicating the local name to assign to the connected printer. &#10; "LPT1"</param>
      /// <param name="remote">String value indicating the name of the remote printer.&#10; "\\\\Server\\Print1"</param>
      /// <param name="updateProfile" optional="true">Optional. Boolean value indicating whether the printer mapping is stored in the current user's profile. If updateProfile is supplied and is true, the mapping is stored in the user profile. The default value is false.</param>
      /// <param name="user" optional="true">Optional. String value indicating the user name. If you are mapping a remote printer using the profile of someone other than current user, you can specify user and password.</param>
      /// <param name="password" optional="true">Optional. String value indicating the user password. If you are mapping a remote printer using the profile of someone other than current user, you can specify user and password.</param>
    }
    this.EnumNetworkDrives = function() {
      /// <summary>Returns the current network drive mapping information.</summary>
      /// <returns type="Collection"/>
      return new Collection
    }
    this.EnumPrinterConnections = function() {
      /// <summary>Returns the current network printer mapping information.</summary>
      /// <returns type="Collection"/>
      return new Collection
    }
    this.MapNetworkDrive = function(local, remote, updateProfile, user, password) {
      /// <summary>Adds a shared network drive to your computer system.</summary>
      /// <param name="local">String value indicating the name by which the mapped drive will be known locally.&#10;"s:"</param>
      /// <param name="remote">String value indicating the share's UNC name (\\xxx\yyy).&#10;"\\\\dns1\\dfs"</param>
      /// <param name="updateProfile" optional="true">Optional. Boolean value indicating whether the mapping information is stored in the current user's profile.</param>
      /// <param name="user" optional="true">Optional. String value indicating the user name. You must supply this argument if you are mapping a network drive using the credentials of someone other than the current user.</param>
      /// <param name="password" optional="true">Optional. String value indicating the user password. You must supply this argument if you are mapping a network drive using the credentials of someone other than the current user.</param>
    }
    this.RemoveNetworkDrive = function(name, force, updateProfile) {
      /// <summary>Removes a shared network drive from your computer system.</summary>
      /// <param name="name">String value indicating the name of the mapped drive you want to remove. The strName parameter can be either a local name or a remote name depending on how the drive is mapped. </param>
      /// <param name="force" optional="true">Optional. Boolean value indicating whether to force the removal of the mapped drive. If bForce is supplied and its value is true, this method removes the connections whether the resource is used or not.</param>
      /// <param name="updateProfile" optional="true">Optional. String value indicating whether to remove the mapping from the user's profile. If updateProfile is supplied and its value is true, this mapping is removed from the user profile. updateProfile is false by default.</param>
    }
    this.RemovePrinterConnection = function(name, force, updateProfile) {
      /// <summary>Removes a shared network printer connection from your computer system.</summary>
      /// <param name="name">String value indicating the name that identifies the printer. It can be a UNC name (in the form \\xxx\yyy) or a local name (such as LPT1).</param>
      /// <param name="force" optional="true">Optional. Boolean value indicating whether to force the removal of the mapped printer. If set to true (the default is false), the printer connection is removed whether or not a user is connected.</param>
      /// <param name="updateProfile" optional="true">Optional. Boolean value. If set to true (the default is false), the change is saved in the user's profile.</param>
    }
    this.SetDefaultPrinter = function(printerName) {
      /// <summary>Assigns a remote printer the role Default Printer.</summary>
      /// <param name="printerName">String value indicating the remote printer's UNC name.&#10;"\\\\research\\library1"</param>
    }
  }

  function Collection() {
    this.Count = function() {
      /// <returns type="Number"/>
    }
    this.Item = function(index) {
      /// <summary>
      /// This collection is an array that associates pairs of items — local names and their associated UNC names. &#10;
      /// Even-numbered items in the collection represent local names.                                             &#10;
      /// Odd-numbered items represent the associated UNC share names.                                             &#10;
      /// The first item in the collection is at index zero (0).                                                   &#10;
      /// </summary>
      /// <returns type="String"/>
    }
  }

}();

+function() {

  ActiveXObject['wscript.shell'] = WScriptShell

  function WScriptShell() {
    /// <field name="CurrentDirectory" type="String">The CurrentDirectory returns a string that contains the fully qualified path of the current working directory of the active process.</field>
    this.CurrentDirectory = ''
    this.AppActivate = function(title) {
      /// <summary>Activates an application window.</summary>
      /// <param name="title" type="String">Specifies which application to activate. This can be a string containing the title of the application (as it appears in the title bar) or the application's Process ID. </param>
      /// <returns type="Boolean"/>
    }
    this.CreateShortcut = function(path) {
      /// <summary>Creates a new shortcut, or opens an existing shortcut.</summary>
      /// <param name="path" type="String">String value indicating the pathname of the shortcut to create.</param>
      return /\.url$/i.test(path)
           ? new WshURLShortcut
           : new WshShortcut
    }
    this.Environment = function(type) {
      /// <summary>Returns the WshEnvironment object (a collection of environment variables).</summary>
      /// <param name="type" type="String">
      /// SYSTEM    &#10;
      /// USER      &#10;
      /// VOLATILE  &#10;
      /// PROCESS   &#10;
      /// </param>
      return function(name) {
        /// <param name="name" type="String">
        /// NUMBER_OF_PROCESSORS                                                   &#10;
        /// PROCESSOR_ARCHITECTURE                                                 &#10;
        /// PROCESSOR_IDENTIFIER                                                   &#10;
        /// PROCESSOR_LEVEL                                                        &#10;
        /// PROCESSOR_REVISION                                                     &#10;
        /// OS                                                                     &#10;
        /// COMSPEC ; Executable file for the command prompt (cmd.exe)             &#10;
        /// HOMEDRIVE ; Primary local drive (C)                                    &#10;
        /// HOMEPATH ; Default directory for users (\users\default)                &#10;
        /// PATH                                                                   &#10;
        /// PATHEXT ; Extensions for executable files (.com, .exe, .bat, or .cmd)  &#10;
        /// PROMPT  ; Command prompt (typically $P$G).                             &#10;
        /// SYSTEMDRIVE                                                            &#10;
        /// SYSTEMROOT                                                             &#10;
        /// WINDIR                                                                 &#10;
        /// TEMP                                                                   &#10;
        /// TMP                                                                    &#10;
        /// </param>
      }
    }
    this.Exec = function(command) {
      /// <summary>Runs an application in a child command-shell, providing access to the StdIn/StdOut/StdErr streams.</summary>
      /// <param name="command" type="String">String value indicating the command line used to run the script. The command line should appear exactly as it would if you typed it at the command prompt.</param>
      /// <returns type="WshScriptExec"/>
      return new WshScriptExec
    }
    this.ExpandEnvironmentStrings = function(environmentVariable) {
      /// <summary>Returns an environment variable's expanded value.</summary>
      /// <param name="environmentVariable" type="String">String value indicating the name of the environment variable you want to expand. </param>
      /// <returns type="String"/>
    }
    this.LogEvent = function(type, message, target) {
      /// <summary>Adds an event entry to a log file.</summary>
      /// <param name="type" type="Number">
      /// 0 SUCCESS         &#10;
      /// 1 ERROR           &#10;
      /// 2 WARNING         &#10;
      /// 4 INFORMATION     &#10;
      /// 8 AUDIT_SUCCESS   &#10;
      /// 16 AUDIT_FAILURE  &#10;
      /// </param>
      /// <param name="message" type="String">String value containing the log entry text. </param>
      /// <param name="target" type="String" optional="true">Optional. String value indicating the name of the computer system where the event log is stored (the default is the local computer system). Applies to Windows NT/2000 only. </param>
      /// <returns type="Boolean"/>
    }
    this.Popup = function(text, wait, title, type) {
      /// <summary>Displays text in a pop-up message box.</summary>
      /// <param name="text" type="String">String value containing the text you want to appear in the pop-up message box.</param>
      /// <param name="wait" type="Number" optional="true">Optional. Numeric value indicating the maximum length of time (in seconds) you want the pop-up message box displayed. </param>
      /// <param name="title" type="String" optional="true">Optional. String value containing the text you want to appear as the title of the pop-up message box. </param>
      /// <param name="type" type="Number" optional="true">Optional. Numeric value indicating the type of buttons and icons you want in the pop-up message box. These determine how the message box is used.
      /// Button Types                             &#10;
      /// 0 Show OK button.                        &#10;
      /// 1 Show OK and Cancel buttons.            &#10;
      /// 2 Show Abort, Retry, and Ignore buttons. &#10;
      /// 3 Show Yes, No, and Cancel buttons.      &#10;
      /// 4 Show Yes and No buttons.               &#10;
      /// 5 Show Retry and Cancel buttons.         &#10;
      /// Icon Types                               &#10;
      /// 16 Show "Stop Mark" icon.                &#10;
      /// 32 Show "Question Mark" icon.            &#10;
      /// 48 Show "Exclamation Mark" icon.         &#10;
      /// 64 Show "Information Mark" icon.         &#10;
      /// </param>
      /// <returns type="Number">Integer value indicating the number of the button the user clicked to dismiss the message box. This is the value returned by the Popup method.
      /// 1 OK button     &#10;
      /// 2 Cancel button &#10;
      /// 3 Abort button  &#10;
      /// 4 Retry button  &#10;
      /// 5 Ignore button &#10;
      /// 6 Yes button    &#10;
      /// 7 No button     &#10;
      /// </returns>
    }
    this.RegDelete = function(name) {
      /// <summary>Deletes a key or one of its values from the registry.</summary>
      /// <param name="name" type="String">
      /// String value indicating the name of the registry key or key value you want to delete.
      /// Specify a key-name by ending "name" with a final backslash; leave it off to specify a value-name.
      /// Fully qualified key-names and value-names are prefixed with a root key.
      /// You may use abbreviated versions of root key names with the RegDelete method.
      /// The five possible root keys you can use are listed in the following table. &#10;
      /// HKCU = HKEY_CURRENT_USER     &#10;
      /// HKLM = HKEY_LOCAL_MACHINE    &#10;
      /// HKCR = HKEY_CLASSES_ROOT     &#10;
      /// HKEY_USERS = (same)          &#10;
      /// HKEY_CURRENT_CONFIG = (same) &#10;
      /// </param>
    }
    this.RegRead = function(name) {
      /// <summary>Returns the value of a key or value-name from the registry.</summary>
      /// <param name="name" type="String">String value indicating the key or value-name whose value you want. &#10;
      /// HKCU = HKEY_CURRENT_USER     &#10;
      /// HKLM = HKEY_LOCAL_MACHINE    &#10;
      /// HKCR = HKEY_CLASSES_ROOT     &#10;
      /// HKEY_USERS = (same)          &#10;
      /// HKEY_CURRENT_CONFIG = (same) &#10;
      /// </param>
      /// <returns type="Variant">
      /// REG_SZ          A string A string                                           &#10;
      /// REG_DWORD       A number An integer                                         &#10;
      /// REG_BINARY      A binary value A VBArray of integers                        &#10;
      /// REG_EXPAND_SZ   An expandable string (e.g., "%windir%\\calc.exe") A string  &#10;
      /// REG_MULTI_SZ    An array of strings A VBArray of strings                    &#10;
      /// </returns>
    }
    this.RegWrite = function(name, value, type) {
      /// <summary>Creates a new key, adds another value-name to an existing key (and assigns it a value), or changes the value of an existing value-name.</summary>
      /// <param name="name" type="String">String value indicating the key-name, value-name, or value you want to create, add, or change. &#10;
      /// HKCU = HKEY_CURRENT_USER     &#10;
      /// HKLM = HKEY_LOCAL_MACHINE    &#10;
      /// HKCR = HKEY_CLASSES_ROOT     &#10;
      /// HKEY_USERS = (same)          &#10;
      /// HKEY_CURRENT_CONFIG = (same) &#10;
      /// </param>
      /// <param name="value" type="String">The name of the new key you want to create, the name of the value you want to add to an existing key, or the new value you want to assign to an existing value-name.</param>
      /// <param name="type" type="String" optional="true">Optional. String value indicating the value's data type. &#10;
      /// REG_SZ = string        &#10;
      /// REG_BINARY = number    &#10;
      /// REG_DWORD  = number    &#10;
      /// REG_EXPAND_SZ = string &#10;
      /// If undefined: &#10;
      /// String = REG_SZ | REG_EXPAND_SZ  &#10;
      /// Integer = REG_DWORD | REG_BINARY &#10;
      /// </param>
    }
    this.Run = function(command, windowStyle, wait) {
      /// <summary>Runs a program in a new process.</summary>
      /// <param name="command" type="String">String value indicating the command line you want to run. You must include any parameters you want to pass to the executable file. &#10;
      /// "notepad "+ WScript.ScriptFullName, 1, true   &#10;
      /// "cmd /K CD C:\ &amp; Dir"                     &#10;
      /// </param>
      /// <param name="windowStyle" type="Number" optional="true">Optional. Integer value indicating the appearance of the program's window. Note that not all programs make use of this information. &#10;
      /// 0  Hides the window and activates another window. &#10;
      /// 1  Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when displaying the window for the first time. &#10;
      /// 2  Activates the window and displays it as a minimized window.  &#10;
      /// 3  Activates the window and displays it as a maximized window.  &#10;
      /// 4  Displays a window in its most recent size and position. The active window remains active. &#10;
      /// 5  Activates the window and displays it in its current size and position. &#10;
      /// 6  Minimizes the specified window and activates the next top-level window in the Z order. &#10;
      /// 7  Displays the window as a minimized window. The active window remains active. &#10;
      /// 8  Displays the window in its current state. The active window remains active. &#10;
      /// 9  Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window. &#10;
      /// 10 Sets the show-state based on the state of the program that started the application. &#10;
      /// </param>
      /// <param name="wait" type="Boolean" optional="true">Optional. Boolean value indicating whether the script should wait for the program to finish executing before continuing to the next statement in your script. If set to true, script execution halts until the program finishes, and Run returns any error code returned by the program. If set to false (the default), the Run method returns immediately after starting the program, automatically returning 0 (not to be interpreted as an error code). </param>
      /// <returns type="Number"/>
    }
    this.SendKeys = function(string) {
      /// <summary>Sends one or more keystrokes to the active window (as if typed on the keyboard).</summary>
      /// <param name="string" type="String">String value indicating the keystroke(s) you want to send. &#10;
      /// BACKSPACE {BACKSPACE}, {BS}, or {BKSP} &#10;
      /// BREAK {BREAK}                          &#10;
      /// CAPS LOCK {CAPSLOCK}                   &#10;
      /// DEL or DELETE {DELETE} or {DEL}        &#10;
      /// DOWN ARROW {DOWN}                      &#10;
      /// END {END}                              &#10;
      /// ENTER {ENTER} or ~                     &#10;
      /// ESC {ESC}                              &#10;
      /// HELP {HELP}                            &#10;
      /// HOME {HOME}                            &#10;
      /// INS or INSERT {INSERT} or {INS}        &#10;
      /// LEFT ARROW {LEFT}                      &#10;
      /// NUM LOCK {NUMLOCK}                     &#10;
      /// PAGE DOWN {PGDN}                       &#10;
      /// PAGE UP {PGUP}                         &#10;
      /// PRINT SCREEN {PRTSC}                   &#10;
      /// RIGHT ARROW {RIGHT}                    &#10;
      /// SCROLL LOCK {SCROLLLOCK}               &#10;
      /// TAB {TAB}                              &#10;
      /// UP ARROW {UP}                          &#10;
      /// F1 to F16 {F1} to {F16}                &#10;
      /// SHIFT +                                &#10;
      /// CTRL ^                                 &#10;
      /// ALT %                                  &#10;
      /// </param>
    }
    this.SpecialFolders = function(specialFolder) {
      /// <summary>Returns a SpecialFolders object (a collection of special folders).</summary>
      /// <param name="specialFolder" type="String">
      /// AllUsersDesktop   &#10;
      /// AllUsersStartMenu &#10;
      /// AllUsersPrograms  &#10;
      /// AllUsersStartup   &#10;
      /// Desktop           &#10;
      /// Favorites         &#10;
      /// Fonts             &#10;
      /// MyDocuments       &#10;
      /// NetHood           &#10;
      /// PrintHood         &#10;
      /// Programs          &#10;
      /// Recent            &#10;
      /// SendTo            &#10;
      /// StartMenu         &#10;
      /// Startup           &#10;
      /// Templates         &#10;
      /// </param>
      /// <returns type="String"/>
    }
  }

  function TextStream() {
    /// <field name="AtEndOfLine" type="Boolean">Returns true if the file pointer is positioned immediately before the end-of-line marker in a TextStream file; false if it is not. Read-only.</field>
    /// <field name="AtEndOfStream" type="Boolean">Returns true if the file pointer is at the end of a TextStream file; false if it is not. Read-only.</field>
    /// <field name="Column" type="Number">Read-only property that returns the column number of the current character position in a TextStream file.</field>
    /// <field name="Line" type="Number">Read-only property that returns the current line number in a TextStream file.</field>
    this.AtEndOfLine = true
    this.AtEndOfStream = true
    this.Column = 0
    this.Line = 0
    this.Close = function() {
      /// <summary>Closes an open TextStream file.</summary>
    }
    this.Read = function(characters) {
      /// <summary>Reads a specified number of characters from a TextStream file and returns the resulting string.</summary>
      /// <param name="characters" type="Number">Required. Number of characters you want to read from the file.</param>
      /// <returns type="String"/>
    }
    this.ReadAll = function() {
      /// <summary>Reads an entire TextStream file and returns the resulting string.</summary>
      /// <returns type="String"/>
    }
    this.ReadLine = function() {
      /// <summary>Reads an entire line (up to, but not including, the newline character) from a TextStream file and returns the resulting string.</summary>
      /// <returns type="String"/>
    }
    this.Skip = function(characters) {
      /// <summary>Skips a specified number of characters when reading a TextStream file.</summary>
      /// <param name="characters" type="Number">Required. Number of characters to skip when reading a file.</param>
    }
    this.SkipLine = function() {
      /// <summary>Skips the next line when reading a TextStream file.</summary>
    }
    this.Write = function(text) {
      /// <summary>Writes a specified string to a TextStream file.</summary>
      /// <param name="text" type="String">Required. The text you want to write to the file.</param>
    }
    this.WriteBlankLines = function(lines) {
      /// <summary>Writes a specified number of newline characters to a TextStream file.</summary>
      /// <param name="lines" type="Number">Required. Number of newline characters you want to write to the file.</param>
    }
    this.WriteLine = function(text) {
      /// <summary>Writes a specified string and newline character to a TextStream file.</summary>
      /// <param name="text" type="String" optional="true">Optional. The text you want to write to the file. If omitted, a newline character is written to the file.</param>
    }
  }

  function WshScriptExec() {
    /// <field name="Status" type="Number">WshRunning ( = 0) The job is still running. WshFinished ( = 1) The job has completed.</field>
    this.Status = 0
    this.StdOut = new TextStream
    this.StdIn = new TextStream
    this.StdErr = new TextStream
    this.Terminate = function() {
      /// <summary>Instructs the script engine to end the process started by the Exec method.</summary>
    }
  }

  function WshShortcut() {
    /// <field name="Arguments" type="String"></field>
    /// <field name="Description" type="String">The Description property contains a string value describing a shortcut.</field>
    /// <field name="FullName" type="String">The FullName property contains a read-only string value indicating the fully qualified path to the shortcut's target.</field>
    /// <field name="Hotkey" type="String">Assigns a key-combination to a shortcut. "ALT+ , CTRL+ , SHIFT+ , EXT+"</field>
    /// <field name="IconLocation" type="String">The string should contain a fully qualified path and an index associated with the icon. "notepad.exe, 0"</field>
    /// <field name="TargetPath" type="String">This property is for the shortcut's target path only. Any arguments to the shortcut must be placed in the Argument's property.</field>
    /// <field name="WindowStyle" type="String">Sets the window style for the program being run. &#10;
    /// 1 Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. &#10;
    /// 3 Activates the window and displays it as a maximized window. &#10;
    /// 7 Minimizes the window and activates the next top-level window. &#10;
    /// </field>
    /// <field name="WorkingDirectory" type="String">Directory in which the shortcut starts. SpecialFolders("Desktop")</field>
    this.Arguments = ''
    this.Description = ''
    this.FullName = ''
    this.Hotkey = ''
    this.IconLocation = ''
    this.TargetPath = ''
    this.WindowStyle = ''
    this.WorkingDirectory = ''
    this.Save = function() {
      /// <summary>Saves a shortcut object to disk.</summary>
    }
  }

  function WshURLShortcut() {
    /// <field name="TargetPath" type="String"></field>
    this.TargetPath = ''
    this.Save = function() {
      /// <summary>Saves a shortcut object to disk.</summary>
    }
  }

}();

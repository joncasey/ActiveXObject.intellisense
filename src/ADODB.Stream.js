///<reference path="ActiveXObject.js">
!function() { ActiveXObject['adodb.stream'] = c
function c() {}
c.prototype = {constructor:c // http://www.w3schools.com/ado/ado_ref_stream.asp
, 'Cancel' : function() {
    ///<summary>Cancels an execution of an Open call on a Stream object</summary>
  }
, 'Close' : function() {
    ///<summary>Closes a Stream object</summary>
  }
, 'CopyTo' : function(stream, chars) {
    ///<summary>Copies a specified number of characters/bytes from one Stream object into another Stream object</summary>
    ///<param name="stream">An object variable value that contains a reference to an open Stream object. The current Stream is copied to the destination Stream specified by DestStream. The destination Stream must already be open. If not, a run-time error occurs.</param>
    ///<param name="chars">Optional. An Integer value that specifies the number of bytes or characters to be copied from the current position in the source Stream to the destination Stream. The default value is –1, which specifies that all characters or bytes are copied from the current position to EOS.</param>
  }
, 'Flush' : function() {
    ///<summary>Sends the contents of the Stream buffer to the associated underlying object</summary>
  }
, 'LoadFromFile' : function(fileName) {
    ///<summary>Loads the contents of a file into a Stream object</summary>
    ///<param name="fileName">A String value that contains the name of a file to be loaded into the Stream. FileName can contain any valid path and name in UNC format. If the specified file does not exist, a run-time error occurs.</param>
  }
, 'Open' : function(source, mode, options, username, password) {
    ///<summary>Opens a Stream object</summary>
    ///<param name="source">Optional. A Variant value that specifies the source of data for the Stream. Source may contain an absolute URL string that points to an existing node in a well-known tree structure, such as an e-mail or file system. A URL should be specified by using the URL keyword ("URL=scheme://server/folder"). Alternatively, Source may contain a reference to an already open Record object, which opens the default stream associated with the Record. If Source is not specified, a Stream is instantiated and opened, associated with no underlying source by default. For more information about URL schemes and their associated providers, see Absolute and Relative URLs.</param>
    ///<param name="mode">Optional. A ConnectModeEnum value that specifies the access mode for the resultant Stream (for example, read/write or read-only). Default value is adModeUnknown. See the Mode property for more information about access modes. If Mode is not specified, it is inherited by the source object. For example, if the source Record is opened in read-only mode, the Stream will also be opened in read-only mode by default.</param>
    ///<param name="options">Optional. A StreamOpenOptionsEnum value. Default value is adOpenStreamUnspecified.</param>
    ///<param name="username">Optional. A String value that contains the user identification that, if it is needed, accesses the Stream object.</param>
    ///<param name="password">Optional. A String value that contains the password that, if it is needed, accesses the Stream object.</param>
  }
, 'Read' : function(bytes) {
    ///<summary>Reads the entire stream or a specified number of bytes from a binary Stream object</summary>
    ///<param name="bytes">Optional. A Long value that specifies the number of bytes to read from the file or the StreamReadEnum value adReadAll, which is the default.</param>
  }
, 'ReadText' : function(chars) {
    ///<summary>Reads the entire stream, a line, or a specified number of characters from a text Stream object</summary>
    ///<param name="chars">Optional. A Long value that specifies the number of characters to read from the file, or a StreamReadEnum value. The default value is adReadAll.
    ///&#10; adReadAll  -1 Default. Reads all bytes from the stream, from the current position onwards to the EOS marker. This is the only valid StreamReadEnum value with binary streams (Type is adTypeBinary).
    ///&#10; adReadLine -2 Reads the next line from the stream (designated by the LineSeparator property).
    ///</param>
    ///<returns type="String"/>
  }
, 'SaveToFile' : function(fileName, options) {
    ///<summary>Saves the binary contents of a Stream object to a file</summary>
    ///<param name="fileName">A String value that contains the fully-qualified name of the file to which the contents of the Stream will be saved. You can save to any valid local location, or any location you have access to via a UNC value.</param>
    ///<param name="options">A SaveOptionsEnum value that specifies whether a new file should be created by SaveToFile, if it does not already exist. Default value is adSaveCreateNotExists. With these options you can specify that an error occurs if the specified file does not exist. You can also specify that SaveToFile overwrites the current contents of an existing file.
    ///&#10; adSaveCreateNotExist  1 Default. Creates a new file if the file specified by the FileName parameter does not already exist.
    ///&#10; adSaveCreateOverWrite 2 Overwrites the file with the data from the currently open Stream object, if the file specified by the Filename parameter already exists. If the file specified by the Filename parameter does not exist, a new file is created.
    ///</param>
  }
, 'SetEOS' : function() {
    ///<summary>Sets the current position to be the end of the stream (EOS)</summary>
  }
, 'SkipLine' : function() {
    ///<summary>Skips a line when reading a text Stream</summary>
  }
, 'Write' : function(buffer) {
    ///<summary>Writes binary data to a binary Stream object</summary>
    ///<param name="buffer">A Variant that contains an array of bytes to be written.</param>
  }
, 'WriteText' : function(data, options) {
    ///<summary>Writes character data to a text Stream object</summary>
    ///<param name="data">A String value that contains the text in characters to be written.</param>
    ///<param name="options">Optional. A StreamWriteEnum value that specifies whether a line separator character must be written at the end of the specified string.
    ///&#10; adWriteChar 0 Default. Writes the specified text string (specified by the Data parameter) to the Stream object.
    ///&#10; adWriteLine 1 Writes a text string and a line separator character to a Stream object. If the LineSeparator property is not defined, then this returns a run-time error.
    ///</param>
  }
, 'CharSet'       : 'Unicode' // Sets or returns a value that specifies into which character set the contents are to be translated. This property is only used with text Stream objects (type is adTypeText)
, 'EOS'           : false     // Returns whether the current position is at the end of the stream or not
, 'LineSeparator' : -1        // Sets or returns the line separator character used in a text Stream object
                              //   adCRLF               -1 Carriage return line feed [Default]
                              //   adLF                 10 Line feed only
                              //   adCR                 13 Carriage return only
, 'Mode'          : 0         // Sets or returns the available permissions for modifying data
                              //   adModeUnknown        0 Permissions have not been set or cannot be determined.
                              //   adModeRead           1 Read-only.
                              //   adModeWrite          2 Write-only.
                              //   adModeReadWrite      3 Read/write.
                              //   adModeShareDenyRead  4 Prevents others from opening a connection with read permissions.
                              //   adModeShareDenyWrite 8 Prevents others from opening a connection with write permissions.
                              //   adModeShareExclusive 12 Prevents others from opening a connection.
                              //   adModeShareDenyNone  16 Allows others to open a connection with any permissions.
                              //   adModeRecursive      0x400000 Used with adModeShareDenyNone, adModeShareDenyWrite, or adModeShareDenyRead to set permissions on all sub-records of the current Record.
, 'Position'      : 0         // Sets or returns the current position (in bytes) from the beginning of a Stream object
, 'Size'          : 0         // Returns the size of an open Stream object
, 'State'         : 0         // Returns a value describing if the Stream object is open or closed
                              //   adStateClosed        0 The object is closed
                              //   adStateOpen          1 The object is open
                              //   adStateConnecting    2 The object is connecting
                              //   adStateExecuting     4 The object is executing a command
                              //   adStateFetching      8 The rows of the object are being retrieved
, 'Type'          : 2         // Sets or returns the type of data in a Stream object
                              //   adTypeBinary         1 Binary data
                              //   adTypeText           2 Text data [Default]
}
}()
var ActiveXObject = ActiveXObject || function (id) {
  ///<param name="id">
  ///ADODB.Stream
  ///&#10;Msxml2.FreeThreadedDOMDocument
  ///&#10;Scripting.FileSystemObject
  ///&#10;Shell.Application
  ///&#10;WScript.Network
  ///&#10;WScript.Shell
  ///</param>
  return new ActiveXObject[String(arguments[0]).toLowerCase()]
}
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
!function() { ActiveXObject['scripting.filesystemobject'] = FileSystemObject

function FileSystemObject() {
  ///<field name="Drives" type="Drives">Returns the Drives collection.</field>
  this.Drives = new Drives()
  this.BuildPath = function(path, name) {
    ///<summary>The BuildPath method inserts an additional path separator between the existing path and the name, only if necessary.</summary>
    ///<param name="path">Required. Existing path to which name is appended. Path can be absolute or relative and need not specify an existing folder.</param>
    ///<param name="name">Required. Name being appended to the existing path.</param>
    ///<returns type="String"/>
  }
  this.CopyFile = function(source, destination, overwrite) {
    ///<summary>Copies one or more files from one location to another.</summary>
    ///<param name="source">Required. Character string file specification, which can include wildcard characters, for one or more files to be copied.</param>
    ///<param name="destination">Required. Character string destination where the file or files from source are to be copied. Wildcard characters are not allowed. </param>
    ///<param name="overwrite">Optional. Boolean value that indicates if existing files are to be overwritten. If true, files are overwritten; if false, they are not. The default is true. Note that CopyFile will fail if destination has the read-only attribute set, regardless of the value of overwrite.</param>
  }
  this.CopyFolder = function(source, destination, overwrite) {
    ///<summary>Recursively copies a folder from one location to another.</summary>
    ///<param name="source">Required. Character string folder specification, which can include wildcard characters, for one or more folders to be copied. </param>
    ///<param name="destination">Required. Character string destination where the folder and subfolders from source are to be copied. Wildcard characters are not allowed.</param>
    ///<param name="overwrite">Optional. Boolean value that indicates if existing folders are to be overwritten. If true, files are overwritten; if false, they are not. The default is true.</param>
  }
  this.CreateFolder = function(foldername) {
    ///<summary>Creates a folder.</summary>
    ///<param name="foldername">Required. String expression that identifies the folder to create. </param>
  }
  this.CreateTextFile = function(filename, overwrite, unicode) {
    ///<summary>Creates a specified file name and returns a TextStream object that can be used to read from or write to the file.</summary>
    ///<param name="filename">Required. String expression that identifies the file to create.</param>
    ///<param name="overwrite">Optional. Boolean value that indicates whether you can overwrite an existing file. The value is true if the file can be overwritten, false if it can't be overwritten. If omitted, existing files are not overwritten. </param>
    ///<param name="unicode">Optional. Boolean value that indicates whether the file is created as a Unicode or ASCII file. The value is true if the file is created as a Unicode file, false if it's created as an ASCII file. If omitted, an ASCII file is assumed. </param>
    ///<returns type="TextStream"/>
    return new TextStream
  }
  this.DeleteFile = function(filespec, force) {
    ///<summary>Deletes a specified file.</summary>
    ///<param name="filespec">Required. The name of the file to delete. The filespec can contain wildcard characters in the last path component.</param>
    ///<param name="force">Optional. Boolean value that is true if files with the read-only attribute set are to be deleted; false (default) if they are not.</param>
  }
  this.DeleteFolder = function(folderspec, force) {
    ///<summary>Deletes a specified folder and its contents.</summary>
    ///<param name="folderspec">Required. The name of the folder to delete. The folderspec can contain wildcard characters in the last path component.</param>
    ///<param name="force">Optional. Boolean value that is true if folders with the read-only attribute set are to be deleted; false (default) if they are not.</param>
  }
  this.DriveExists = function(drivespec) {
    ///<summary>Returns True if the specified drive exists; False if it does not.&#10;For drives with removable media, the DriveExists method returns true even if there are no media present. Use the IsReady property of the Drive object to determine if a drive is ready.</summary>
    ///<param name="drivespec">Required. A drive letter or a complete path specification. </param>
    ///<returns type="Boolean"/>
  }
  this.FileExists = function(filespec) {
    ///<summary>Returns True if a specified file exists; False if it does not.</summary>
    ///<param name="filespec">Required. The name of the file whose existence is to be determined. A complete path specification (either absolute or relative) must be provided if the file isn't expected to exist in the current folder.</param>
    ///<returns type="Boolean"/>
  }
  this.FolderExists = function(folderspec) {
    ///<summary>Returns True if a specified folder exists; False if it does not.</summary>
    ///<param name="folderspec">Required. The name of the folder whose existence is to be determined. A complete path specification (either absolute or relative) must be provided if the folder isn't expected to exist in the current folder.</param>
    ///<returns type="Boolean"/>
  }
  this.GetAbsolutePathName = function(pathspec) {
    ///<summary>Returns a complete and unambiguous path from a provided path specification.</summary>
    ///<param name="pathspec">Required. Path specification to change to a complete and unambiguous path.
    ///&#10;"c:".........................."c:\mydocuments\reports" 
    ///&#10;"c:.."........................"c:\mydocuments" 
    ///&#10;"c:\\"........................"c:\" 
    ///&#10;"c:*.*\\may97"................"c:\mydocuments\reports\*.*\may97" 
    ///&#10;"region1"....................."c:\mydocuments\reports\region1" 
    ///&#10;"c:\\..\\..\\mydocuments"....."c:\mydocuments" 
    ///</param>
    ///<returns type="String"/>
  }
  this.GetBaseName = function(path) {
    ///<summary>Returns a string containing the base name of the last component, less any file extension, in a path.</summary>
    ///<param name="path">Required. The path specification for the component whose base name is to be returned.</param>
    ///<returns type="String"/>
  }
  this.GetDrive = function(drivespec) {
    ///<summary>Returns a Drive object corresponding to the drive in a specified path.</summary>
    ///<param name="drivespec">Required. The drivespec argument can be a drive letter (c), a drive letter with a colon appended (c:), a drive letter with a colon and path separator appended (c:\), or any network share specification (\\computer2\share1).</param>
    ///<returns type="Drive"/>
    return new Drive
  }
  this.GetDriveName = function(path) {
    ///<summary>Returns a string containing the name of the drive for a specified path.</summary>
    ///<param name="path">Required. The path specification for the component whose drive name is to be returned</param>
    ///<returns type="String"/>
  }
  this.GetExtensionName = function(path) {
    ///<summary>Returns a string containing the extension name for the last component in a path.</summary>
    ///<param name="path">Required. The path specification for the component whose extension name is to be returned.</param>
    ///<returns type="String"/>
  }
  this.GetFile = function(filespec) {
    ///<summary>Returns a File object corresponding to the file in a specified path.</summary>
    ///<param name="filespec">Required. The filespec is the path (absolute or relative) to a specific file.</param>
    ///<returns type="File"/>
    return new File
  }
  this.GetFileName = function(pathspec) {
    ///<summary>Returns the last component of specified path that is not part of the drive specification.</summary>
    ///<param name="pathspec">Required. The path (absolute or relative) to a specific file.</param>
    ///<returns type="String"/>
  }
  this.GetFolder = function(folderspec) {
    ///<summary>Returns a Folder object corresponding to the folder in a specified path.</summary>
    ///<param name="folderspec">Required. The folderspec is the path (absolute or relative) to a specific folder.</param>
    ///<returns type="Folder"/>
    return new Folder
  }
  this.GetParentFolderName = function(path) {
    ///<summary>Returns a string containing the name of the parent folder of the last component in a specified path.</summary>
    ///<param name="path">Required. The path specification for the component whose parent folder name is to be returned.</param>
    ///<returns type="String"/>
  }
  this.GetSpecialFolder = function(folderspec) {
    ///<summary>Returns the special folder object specified.</summary>
    ///<param name="folderspec">Required. The name of the special folder to be returned. Can be any of the following.
    ///&#10;0 WindowsFolder, The Windows folder contains files installed by the Windows operating system. 
    ///&#10;1 SystemFolder, The System folder contains libraries, fonts, and device drivers. 
    ///&#10;2 TemporaryFolder, The Temp folder is used to store temporary files. Its path is found in the TMP environment variable. 
    ///</param>
    ///<returns type="Folder"/>
    return new Folder
  }
  this.GetTempName = function() {
    ///<summary>Returns a randomly generated temporary file or folder name that is useful for performing operations that require a temporary file or folder.</summary>
    ///<returns type="String"/>
  }
  this.MoveFile = function(source, destination) {
    ///<summary>Moves one or more files from one location to another.</summary>
    ///<param name="source">Required. The path to the file or files to be moved. The source argument string can contain wildcard characters in the last path component only. </param>
    ///<param name="destination">Required. The path where the file or files are to be moved. The destination argument can't contain wildcard characters. </param>
  }
  this.MoveFolder = function(source, destination) {
    ///<summary>Moves one or more folders from one location to another.</summary>
    ///<param name="source">Required. The path to the folder or folders to be moved. The source argument string can contain wildcard characters in the last path component only.</param>
    ///<param name="destination">Required. The path where the folder or folders are to be moved. The destination argument can't contain wildcard characters.</param>
  }
  this.OpenTextFile = function(filename, iomode, create, format) {
    ///<summary>Opens a specified file and returns a TextStream object that can be used to read from, write to, or append to the file.</summary>
    ///<param name="filename">Required. String expression that identifies the file to open.</param>
    ///<param name="iomode">Optional. Can be one of three constants:
    ///&#10;1 ForReading, Open a file for reading only. You can't write to this file. 
    ///&#10;2 ForWriting, Open a file for writing. If a file with the same name exists, its previous contents are overwritten.
    ///&#10;8 ForAppending, Open a file and write to the end of the file. 
    ///</param>
    ///<param name="create">Optional. Boolean value that indicates whether a new file can be created if the specified filename doesn't exist. The value is True if a new file is created, False if it isn't created. If omitted, a new file isn't created.</param>
    ///<param name="format">Optional. One of three Tristate values used to indicate the format of the opened file. If omitted, the file is opened as ASCII.
    ///&#10; 0 TristateFalse, Open the file as ASCII. 
    ///&#10;-1 TristateTrue, Open the file as Unicode. 
    ///&#10;-2 TristateUseDefault, Open the file using the system default. 
    ///</param>
    ///<returns type="TextStream"/>
    return new TextStream
  }
}

function Drive() {
  ///<field name="AvailableSpace" type="Number">Returns the amount of space available to a user on the specified drive or network share.</field>
  ///<field name="DriveLetter" type="String">Returns the drive letter of a physical local drive or a network share. Read-only.</field>
  ///<field name="DriveType" type="Number">Returns a value indicating the type of a specified drive.&#10;0  Unknown&#10;1  Removable&#10;2  Fixed&#10;3  Network&#10;4  CD ROM&#10;5  RAM Disk</field>
  ///<field name="FileSystem" type="String">Returns the type of file system in use for the specified drive. FAT, NTFS, CDFS</field>
  ///<field name="FreeSpace" type="Number">Returns the amount of free space available to a user on the specified drive or network share. Read-only.</field>
  ///<field name="IsReady" type="Boolean">Returns True if the specified drive is ready; False if it is not.</field>
  ///<field name="Path" type="String">Returns the path for a specified file, folder, or drive.</field>
  ///<field name="RootFolder" type="Folder">Returns a Folder object representing the root folder of a specified drive. Read-only</field>
  ///<field name="SerialNumber" type="Number">Returns the decimal serial number used to uniquely identify a disk volume.</field>
  ///<field name="ShareName" type="String">Returns the network share name for a specified drive.</field>
  ///<field name="TotalSize" type="Number">Returns the total space, in bytes, of a drive or network share.</field>
  ///<field name="VolumeName" type="String">Sets or returns the volume name of the specified drive. Read/write.</field>
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
  ///<field name="Attributes" type="Number">Sets or returns the attributes of files or folders. Read/write or read-only, depending on the attribute.
  ///&#10;0 Normal file. No attributes are set.
  ///&#10;1 Read-only file. Attribute is read/write.
  ///&#10;2 Hidden file. Attribute is read/write.
  ///&#10;4 System file. Attribute is read/write.
  ///&#10;8 Disk drive volume label. Attribute is read-only.
  ///&#10;16 Directory or Folder. Attribute is read-only.
  ///&#10;32 Archive, File has changed since last backup. Attribute is read/write.
  ///&#10;64 Alias, Link or shortcut. Attribute is read-only.
  ///&#10;128 Compressed file. Attribute is read-only.
  ///</field>
  ///<field name="DateCreated" type="String">Read-only.</field>
  ///<field name="DateLastAccessed" type="String">Read-only.</field>
  ///<field name="DateLastModified" type="String">Read-only.</field>
  ///<field name="Drive" type="String">Returns the drive letter of the drive on which the specified file or folder resides. Read-only.</field>
  ///<field name="Name" type="String">Sets or returns the name of a specified file or folder. Read/write.</field>
  ///<field name="ParentFolder" type="">Returns the folder object for the parent of the specified file or folder. Read-only.</field>
  ///<field name="Path" type="String">Returns the path for a specified file, folder, or drive.</field>
  ///<field name="ShortName" type="String">Returns the short name used by programs that require the earlier 8.3 naming convention.</field>
  ///<field name="ShortPath" type="String">Returns the short path used by programs that require the earlier 8.3 file naming convention.</field>
  ///<field name="Size" type="">For files, returns the size, in bytes, of the specified file. For folders, returns the size, in bytes, of all files and subfolders contained in the folder.</field>
  ///<field name="Type" type="String">Returns information about the type of a file or folder. For example, for files ending in .TXT, "Text Document" is returned.</field>
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
    ///<summary>Copies a specified file or folder from one location to another.</summary>
    ///<param name="destination">Destination where the file or folder is to be copied. Wildcard characters are not allowed.</param>
    ///<param name="overwrite">Optional. Boolean value that is True (default) if existing file is to be overwritten; False if they are not.</param>
  }
  this.Delete = function(force) {
    ///<summary>Deletes a specified file.</summary>
    ///<param name="force">Optional. Boolean value that is True if file with the read-only attribute set is to be deleted; False (default) if it is not.</param>
  }
  this.Move = function(destination) {
    ///<summary>Moves a specified file from one location to another.</summary>
    ///<param name="destination">Required. Destination where the file or folder is to be moved. Wildcard characters are not allowed.</param>
  }
  this.OpenAsTextStream = function(iomode, format) {
    ///<summary>Opens a specified file and returns a TextStream object that can be used to read from, write to, or append to the file.</summary>
    ///<param name="iomode">Optional. Indicates input/output mode. Can be one of three constants:
    ///&#10;1 ForReading, Open a file for reading only. You can't write to this file. 
    ///&#10;2 ForWriting, Open a file for writing. If a file with the same name exists, its previous contents are overwritten.
    ///&#10;8 ForAppending, Open a file and write to the end of the file. 
    ///</param>
    ///<param name="format">Optional. One of three Tristate values used to indicate the format of the opened file. If omitted, the file is opened as ASCII.
    ///&#10; 0 TristateFalse, Open the file as ASCII. 
    ///&#10;-1 TristateTrue, Open the file as Unicode. 
    ///&#10;-2 TristateUseDefault, Open the file using the system default. 
    ///</param>
    ///<returns type="TextStream"/>
    return new TextStream
  }
}

function Folder(parentFolder) {
  ///<field name="Attributes"       type="Number">Sets or returns the attributes of files or folders. Read/write or read-only, depending on the attribute.
  ///&#10;0 Normal file. No attributes are set.
  ///&#10;1 Read-only file. Attribute is read/write.
  ///&#10;2 Hidden file. Attribute is read/write.
  ///&#10;4 System file. Attribute is read/write.
  ///&#10;8 Disk drive volume label. Attribute is read-only.
  ///&#10;16 Directory or Folder. Attribute is read-only.
  ///&#10;32 Archive, File has changed since last backup. Attribute is read/write.
  ///&#10;64 Alias, Link or shortcut. Attribute is read-only.
  ///&#10;128 Compressed file. Attribute is read-only.
  ///</field>
  ///<field name="DateCreated"      type="String">Read-only.</field>
  ///<field name="DateLastAccessed" type="String">Read-only.</field>
  ///<field name="DateLastModified" type="String">Read-only.</field>
  ///<field name="Drive"            type="String">Returns the drive letter of the drive on which the specified file or folder resides. Read-only.</field>
  ///<field name="Files"            type="Files">Returns a Files collection consisting of all File objects contained in the specified folder, including those with hidden and system file attributes set.</field>
  ///<field name="IsRootFolder"     type="Boolean">Returns True if the specified folder is the root folder; False if it is not.</field>
  ///<field name="Name"             type="String">Sets or returns the name of a specified file or folder. Read/write.</field>
  ///<field name="ParentFolder"     type="Folder">Returns the folder object for the parent of the specified file or folder. Read-only.</field>
  ///<field name="Path"             type="String">Returns the path for a specified file, folder, or drive.</field>
  ///<field name="ShortName"        type="String">Returns the short name used by programs that require the earlier 8.3 naming convention.</field>
  ///<field name="ShortPath"        type="String">Returns the short path used by programs that require the earlier 8.3 file naming convention.</field>
  ///<field name="Size"             type="Number">For files, returns the size, in bytes, of the specified file. For folders, returns the size, in bytes, of all files and subfolders contained in the folder.</field>
  ///<field name="SubFolders"       type="Folders">Returns a Folders collection consisting of all folders contained in a specified folder, including those with hidden and system file attributes set.</field>
  ///<field name="Type"             type="String">Returns information about the type of a file or folder. For example, for files ending in .TXT, "Text Document" is returned.</field>
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
    ///<summary>Copies a specified file or folder from one location to another.</summary>
    ///<param name="destination">Required. Destination where the file or folder is to be copied. Wildcard characters are not allowed. </param>
    ///<param name="overwrite">Optional. Boolean value that is True (default) if existing files or folders are to be overwritten; False if they are not. </param>
  }
  this.Delete = function(force) {
    ///<summary>Deletes a specified file or folder.</summary>
    ///<param name="force">Optional. Boolean value that is True if files or folders with the read-only attribute set are to be deleted; False (default) if they are not. </param>
  }
  this.Move = function(destination) {
    ///<summary>Moves a specified file or folder from one location to another.</summary>
    ///<param name="destination">Required. Destination where the file or folder is to be moved. Wildcard characters are not allowed. </param>
  }
}

function TextStream() {
  ///<field name="AtEndOfLine" type="Boolean">Returns true if the file pointer is positioned immediately before the end-of-line marker in a TextStream file; false if it is not. Read-only.</field>
  ///<field name="AtEndOfStream" type="Boolean">Returns true if the file pointer is at the end of a TextStream file; false if it is not. Read-only.</field>
  ///<field name="Column" type="Number">Read-only property that returns the column number of the current character position in a TextStream file.</field>
  ///<field name="Line" type="Number">Read-only property that returns the current line number in a TextStream file.</field>
  this.AtEndOfLine   = true
  this.AtEndOfStream = true
  this.Column        = 0
  this.Line          = 0
  this.Close = function() {
    ///<summary>Closes an open TextStream file.</summary>
  }
  this.Read = function(characters) {
    ///<summary>Reads a specified number of characters from a TextStream file and returns the resulting string.</summary>
    ///<param name="characters">Required. Number of characters you want to read from the file.</param>
    ///<returns type="String"/>
  }
  this.ReadAll = function() {
    ///<summary>Reads an entire TextStream file and returns the resulting string.</summary>
    ///<returns type="String"/>
  }
  this.ReadLine = function() {
    ///<summary>Reads an entire line (up to, but not including, the newline character) from a TextStream file and returns the resulting string.</summary>
    ///<returns type="String"/>
  }
  this.Skip = function(characters) {
    ///<summary>Skips a specified number of characters when reading a TextStream file.</summary>
    ///<param name="characters">Required. Number of characters to skip when reading a file.</param>
  }
  this.SkipLine = function() {
    ///<summary>Skips the next line when reading a TextStream file.</summary>
  }
  this.Write = function(string) {
    ///<summary>Writes a specified string to a TextStream file.</summary>
    ///<param name="string">Required. The text you want to write to the file.</param>
  }
  this.WriteBlankLines = function(lines) {
    ///<summary>Writes a specified number of newline characters to a TextStream file.</summary>
    ///<param name="lines">Required. Number of newline characters you want to write to the file.</param>
  }
  this.WriteLine = function(string) {
    ///<summary>Writes a specified string and newline character to a TextStream file.</summary>
    ///<param name="string">Optional. The text you want to write to the file. If omitted, a newline character is written to the file.</param>
  }
}

function Drives() {
  this.Count = 0
  this.Item = function(key) {
    ///<summary>Returns an item based on the specified key.</summary>
    ///<returns type="Drive"/>
    return new Drive
  }
}

function Files() {
  this.Count = 0
  this.Item = function(key) {
    ///<summary>Returns an item based on the specified key.</summary>
    ///<returns type="File"/>
    return new File
  }
}

function Folders() {
  this.Count = 0
  this.Add = function(foldername) {
    ///<summary>Adds a new folder to a Folders collection.</summary>
    ///<param name="foldername">Required. The name of the new Folder being added.</param>
  }
  this.Item = function(key) {
    ///<summary>Returns an item based on the specified key.</summary>
    ///<returns type="Folder"/>
    return new Folder
  }
}

}()
!function() { ActiveXObject['shell.application'] = c
function c() {}
c.prototype = {constructor:c
, 'AddToRecent' : function(file, category) {
    ///<summary>Adds a file to the most recently used (MRU) list.</summary>
    ///<param name="file">string, path of the file to add to the list of recently used documents&#10;null, to clear the recent documents folder</param>
    ///<param name="category">string, name of the category in which ot place the file (optional)</param>
  }
, 'BrowseForFolder' : function(hwnd, title, options, rootFolder) {
    ///<summary>Creates a dialog box that enables the user to select a folder and then returns the selected folder's Folder object.</summary>
    ///<param name="hwnd">number, handle of the parent window of the dialog box, or zero.</param>
    ///<param name="title">string, represents the title displayed inside the Browse dialog box.</param>
    ///<param name="options">number, zero or a combination of the ulFlags in BROWSEINFO</param>
    ///<param name="rootFolder">user can not browse higher than this folder
    ///&#10;string, specifying a path
    ///&#10;number, one of the ShellSpecialFolderConstants (ssf)
    ///&#10;undefined, will use DESKTOP</param>
    ///<returns type="Folder"/>
    return new Folder
  }
, 'CanStartStopService' : function() {
    ///<summary>Determines if the current user can start and stop the named service.</summary>
    ///<param name=""></param>
  }
, 'CascadeWindows' : function() {
    ///<summary>Cascades all of the windows on the desktop. This method has the same effect as right-clicking the taskbar and selecting Cascade Windows.</summary>
    ///<param name=""></param>
  }
, 'ControlPanelItem' : function() {
    ///<summary>Runs the specified Control Panel (*.cpl) application. If the application is already open, it will activate the running instance.</summary>
    ///<param name=""></param>
  }
, 'EjectPC' : function() {
    ///<summary>Ejects the computer from its docking station. This is the same as clicking the Start menu and selecting Eject PC, if your computer supports this command.</summary>
    ///<param name=""></param>
  }
, 'Explore' : function() {
    ///<summary>Opens a specified folder in a Windows Explorer window.</summary>
    ///<param name=""></param>
  }
, 'ExplorerPolicy' : function() {
    ///<summary>Gets the value for a specified Internet Explorer policy.</summary>
    ///<param name=""></param>
  }
, 'FileRun' : function() {
    ///<summary>Displays the Run dialog to the user. This method has the same effect as clicking the Start menu and selecting Run.</summary>
    ///<param name=""></param>
  }
, 'FindComputer' : function() {
    ///<summary>Displays the Search Results: Computers dialog box. The dialog box shows the result of the search for a specified computer.</summary>
    ///<param name=""></param>
  }
, 'FindFiles' : function() {
    ///<summary>Displays the Find: All Files dialog box. This is the same as clicking the Start menu and then selecting Search (or its equivalent under systems earlier than Windows XP.</summary>
    ///<param name=""></param>
  }
, 'FindPrinter' : function() {
    ///<summary>Displays the Find Printer dialog box.</summary>
    ///<param name=""></param>
  }
, 'GetSetting' : function() {
    ///<summary>Retrieves a global Shell setting.</summary>
    ///<param name=""></param>
  }
, 'GetSystemInformation' : function() {
    ///<summary>Retrieves system information.</summary>
    ///<param name=""></param>
  }
, 'Help' : function() {
    ///<summary>Displays the Windows Help and Support Center. This method has the same effect as clicking the Start menu and selecting Help and Support.</summary>
    ///<param name=""></param>
  }
, 'IsRestricted' : function() {
    ///<summary>Retrieves a group's restriction setting from the registry.</summary>
    ///<param name=""></param>
  }
, 'IsServiceRunning' : function() {
    ///<summary>Returns a value that indicates whether a particular service is running.</summary>
    ///<param name=""></param>
  }
, 'MinimizeAll' : function() {
    ///<summary>Minimizes all of the windows on the desktop. This method has the same effect as right-clicking the taskbar and selecting Minimize All Windows on older systems or clicking the Show Desktop icon in the Quick Launch area of the taskbar in Windows 2000 or Windows XP.</summary>
    ///<param name=""></param>
  }
, 'NameSpace' : function(dir) {
    ///<summary>Creates and returns a Folder object for the specified folder.</summary>
    ///<param name="dir">
    ///&#10;string, path of the folder
    ///&#10;number, a value from the ShellSpecialFolderConstants (ssf)
    ///</param>
    return new Folder
  }
, 'Open' : function() {
    ///<summary>Opens the specified folder.</summary>
    ///<param name=""></param>
  }
, 'RefreshMenu' : function() {
    ///<summary>Refreshes the contents of the Start menu. Used only with systems preceding Windows XP.</summary>
    ///<param name=""></param>
  }
, 'ServiceStart' : function() {
    ///<summary>Starts a named service.</summary>
    ///<param name=""></param>
  }
, 'ServiceStop' : function() {
    ///<summary>Stops a named service.</summary>
    ///<param name=""></param>
  }
, 'SetTime' : function() {
    ///<summary>Displays the Date and Time Properties dialog box. This method has the same effect as right-clicking the clock in the taskbar status area and selecting Adjust Date/Time.</summary>
    ///<param name=""></param>
  }
, 'ShellExecute' : function() {
    ///<summary>Performs a specified operation on a specified file.</summary>
    ///<param name=""></param>
  }
, 'ShowBrowserBar' : function() {
    ///<summary>Displays a browser bar.</summary>
    ///<param name=""></param>
  }
, 'ShutdownWindows' : function() {
    ///<summary>Displays the Shut Down Windows dialog box. This is the same as clicking the Start menu and selecting Shut Down.</summary>
    ///<param name=""></param>
  }
, 'Suspend' : function() {
    ///<summary></summary>
    ///<param name=""></param>
  }
, 'TileHorizontally' : function() {
    ///<summary>Tiles all of the windows on the desktop horizontally. This method has the same effect as right-clicking the taskbar and selecting Tile Windows Horizontally.</summary>
    ///<param name=""></param>
  }
, 'TileVertically' : function() {
    ///<summary>Tiles all of the windows on the desktop vertically. This method has the same effect as right-clicking the taskbar and selecting Tile Windows Vertically.</summary>
    ///<param name=""></param>
  }
, 'TrayProperties' : function() {
    ///<summary>Displays the Taskbar and Start Menu Properties dialog box. This method has the same effect as right-clicking the taskbar and selecting Properties.</summary>
    ///<param name=""></param>
  }
, 'UndoMinimizeALL' : function() {
    ///<summary>Restores all desktop windows to the same state they were in before the last MinimizeAll command. This method has the same effect as right-clicking the taskbar and selecting Undo Minimize All Windows on older systems or a second clicking of the Show Desktop icon in the Quick Launch area of the taskbar in Windows 2000 or Windows XP.</summary>
    ///<param name=""></param>
  }
, 'Windows' : function() {
    ///<summary>Creates and returns a ShellWindows object. This object represents a collection of all of the open windows that belong to the Shell.</summary>
    ///<param name=""></param>
  }
, 'WindowsSecurity' : function() {
    ///<summary>Displays the Windows Security dialog box.</summary>
    ///<param name=""></param>
  }
, 'WindowSwitcher' : function() {
    ///<summary>Displays your open windows in a 3D stack that you can flip through.</summary>
    ///<param name=""></param>
  }
, ulFlags:{__namespace:true
  , BIF_RETURNONLYFSDIRS    :     1 // Only return file system directories. If the user selects folders that are not part of the file system, the OK button is grayed. Note  The OK button remains enabled for "\\server" items, as well as "\\server\share" and directory items. However, if the user selects a "\\server" item, passing the PIDL returned by SHBrowseForFolder to SHGetPathFromIDList fails.
  , BIF_DONTGOBELOWDOMAIN   :     2 // Do not include network folders below the domain level in the dialog box's tree view control.
  , BIF_STATUSTEXT          :     4 // Include a status area in the dialog box. The callback function can set the status text by sending messages to the dialog box. This flag is not supported when BIF_NEWDIALOGSTYLE is specified.
  , BIF_RETURNFSANCESTORS   :     8 // Only return file system ancestors. An ancestor is a subfolder that is beneath the root folder in the namespace hierarchy. If the user selects an ancestor of the root folder that is not part of the file system, the OK button is grayed.
  , BIF_EDITBOX             :    10 // Version 4.71. Include an edit control in the browse dialog box that allows the user to type the name of an item.
  , BIF_VALIDATE            :    20 // Version 4.71. If the user types an invalid name into the edit box, the browse dialog box calls the application's BrowseCallbackProc with the BFFM_VALIDATEFAILED message. This flag is ignored if BIF_EDITBOX is not specified.
  , BIF_NEWDIALOGSTYLE      :    40 // Version 5.0. Use the new user interface. Setting this flag provides the user with a larger dialog box that can be resized. The dialog box has several new capabilities, including: drag-and-drop capability within the dialog box, reordering, shortcut menus, new folders, delete, and other shortcut menu commands. Note  If COM is initialized through CoInitializeEx with the COINIT_MULTITHREADED flag set, SHBrowseForFolder fails if BIF_NEWDIALOGSTYLE is passed.
  , BIF_BROWSEINCLUDEURLS   :    80 // Version 5.0. The browse dialog box can display URLs. The BIF_USENEWUI and BIF_BROWSEINCLUDEFILES flags must also be set. If any of these three flags are not set, the browser dialog box rejects URLs. Even when these flags are set, the browse dialog box displays URLs only if the folder that contains the selected item supports URLs. When the folder's IShellFolder::GetAttributesOf method is called to request the selected item's attributes, the folder must set the SFGAO_FOLDER attribute flag. Otherwise, the browse dialog box will not display the URL.
  , BIF_USENEWUI            :    50 // Version 5.0. Use the new user interface, including an edit box. This flag is equivalent to BIF_EDITBOX | BIF_NEWDIALOGSTYLE. Note  If COM is initialized through CoInitializeEx with the COINIT_MULTITHREADED flag set, SHBrowseForFolder fails if BIF_USENEWUI is passed.
  , BIF_UAHINT              :   100 // Version 6.0. When combined with BIF_NEWDIALOGSTYLE, adds a usage hint to the dialog box, in place of the edit box. BIF_EDITBOX overrides this flag.
  , BIF_NONEWFOLDERBUTTON   :   200 // Version 6.0. Do not include the New Folder button in the browse dialog box.
  , BIF_NOTRANSLATETARGETS  :   400 // Version 6.0. When the selected item is a shortcut, return the PIDL of the shortcut itself rather than its target.
  , BIF_BROWSEFORCOMPUTER   :  1000 // Only return computers. If the user selects anything other than a computer, the OK button is grayed.
  , BIF_BROWSEFORPRINTER    :  2000 // Only allow the selection of printers. If the user selects anything other than a printer, the OK button is grayed. In Windows XP and later systems, the best practice is to use a Windows XP-style dialog, setting the root of the dialog to the Printers and Faxes folder (CSIDL_PRINTERS).
  , BIF_BROWSEINCLUDEFILES  :  4000 // Version 4.71. The browse dialog box displays files as well as folders.
  , BIF_SHAREABLE           :  8000 // Version 5.0. The browse dialog box can display sharable resources on remote systems. This is intended for applications that want to expose remote shares on a local system. The BIF_NEWDIALOGSTYLE flag must also be set.
  , BIF_BROWSEFILEJUNCTIONS : 10000 // Windows 7 and later. Allow folder junctions such as a library or a compressed file with a .zip file name extension to be browsed.
  }
, ssf:{__namespace:true
  , ALTSTARTUP         : 29 // File system directory that corresponds to the user's non-localized Startup program group.
  , APPDATA            : 26 // File system directory that serves as a common repository for application-specific data. A typical path is C:\Documents and Settings\username\Application Data.
  , BITBUCKET          : 10 // Virtual folder that contains the objects in the user's Recycle Bin.
  , COMMONALTSTARTUP   : 30 // File system directory that corresponds to the non-localized Startup program group for all users. Valid only for Windows NT systems.
  , COMMONAPPDATA      : 35 // Application data for all users. A typical path is C:\Documents and Settings\All Users\Application Data.
  , COMMONDESKTOPDIR   : 25 // File system directory that contains files and folders that appear on the desktop for all users. A typical path is C:\Documents and Settings\All Users\Desktop. Valid only for Windows NT systems.
  , COMMONFAVORITES    : 31 // File system directory that serves as a common repository for the favorite URLs shared by all users. Valid only for Windows NT systems.
  , COMMONPROGRAMS     : 23 // File system directory that contains the directories for the common program groups that appear on the Start menu for all users. A typical path is C:\Documents and Settings\All Users\Start Menu\Programs. Valid only for Windows NT systems.
  , COMMONSTARTMENU    : 22 // File system directory that contains the programs and folders that appear on the Start menu for all users. A typical path is C:\Documents and Settings\All Users\Start Menu. Valid only for Windows NT systems.
  , COMMONSTARTUP      : 24 // File system directory that contains the programs that appear in the Startup folder for all users. A typical path is C:\Documents and Settings\All Users\Microsoft\Windows\Start Menu\Programs\StartUp. Valid only for Windows NT systems.
  , CONTROLS           :  3 // Virtual folder that contains icons for the Control Panel applications.
  , COOKIES            : 33 // File system directory that serves as a common repository for Internet cookies. A typical path is C:\Documents and Settings\username\Application Data\Microsoft\Windows\Cookies.
  , DESKTOP            :  0 // Windows desktop—the virtual folder that is the root of the namespace.
  , DESKTOPDIRECTORY   : 16 // File system directory used to physically store the file objects that are displayed on the desktop. It is not to be confused with the desktop folder itself, which is a virtual folder. A typical path is C:\Documents and Settings\username\Desktop.
  , DRIVES             : 17 // My Computer—the virtual folder that contains everything on the local computer: storage devices, printers, and Control Panel. This folder can also contain mapped network drives.
  , FAVORITES          :  6 // File system directory that serves as a common repository for the user's favorite URLs. A typical path is C:\Documents and Settings\username\Favorites.
  , FONTS              : 20 // Virtual folder that contains installed fonts. A typical path is C:\Windows\Fonts.
  , HISTORY            : 34 // File system directory that serves as a common repository for Internet history items.
  , INTERNETCACHE      : 32 // File system directory that serves as a common repository for temporary Internet files. A typical path is C:\Users\username\AppData\Local\Microsoft\Windows\Temporary Internet Files.
  , LOCALAPPDATA       : 28 // File system directory that serves as a data repository for local (non-roaming) applications. A typical path is C:\Users\username\AppData\Local.
  , MYPICTURES         : 39 // My Pictures folder. A typical path is C:\Users\username\Pictures.
  , NETHOOD            : 19 // A file system folder that contains any link objects in the My Network Places virtual folder. It is not the same as ssfNETWORK, which represents the network namespace root. A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Network Shortcuts.
  , NETWORK            : 18 // Network Neighborhood—the virtual folder that represents the root of the network namespace hierarchy.
  , PERSONAL           :  5 // File system directory that serves as a common repository for a user's documents. A typical path is C:\Users\username\Documents.
  , PRINTERS           :  4 // Virtual folder that contains installed printers.
  , PRINTHOOD          : 27 // File system directory that contains any link objects in the Printers virtual folder. A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Printer Shortcuts.
  , PROFILE            : 40 // User's profile folder.
  , PROGRAMFILES       : 38 // Program Files folder. A typical path is C:\Program Files.
  , PROGRAMFILESx86    : 48 // Program Files folder. A typical path is C:\Program Files, or C:\Program Files (X86) on a 64-bit computer.
  , PROGRAMS           :  2 // File system directory that contains the user's program groups (which are also file system directories). A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs.
  , RECENT             :  8 // File system directory that contains the user's most recently used documents. A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Recent.
  , SENDTO             :  9 // File system directory that contains Send To menu items. A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\SendTo.
  , STARTMENU          : 11 // File system directory that contains Start menu items. A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Start Menu.
  , STARTUP            :  7 // File system directory that corresponds to the user's Startup program group. The system starts these programs whenever any user first logs into their profile after a reboot. A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\StartUp.
  , SYSTEM             : 37 // The System folder. A typical path is C:\Windows\System32.
  , SYSTEMx86          : 41 // System folder. A typical path is C:\Windows\System32, or C:\Windows\Syswow32 on a 64-bit computer.
  , TEMPLATES          : 21 // File system directory that serves as a common repository for document templates.
  , WINDOWS            : 36 // Windows directory. This corresponds to the %windir% or %SystemRoot% environment variables. A typical path is C:\Windows.
  }
}

// Folder                                                         // http://msdn.microsoft.com/en-us/library/bb787868%28v=vs.85%29.aspx

function Folder() {}
Folder.prototype = {constructor:Folder
, 'CopyHere' : function() {
    ///<summary>Copies an item or items to a folder.</summary>
    ///<param name=""></param>
  }
, 'GetDetailsOf' : function() {
    ///<summary>Retrieves details about an item in a folder. For example, its size, type, or the time of its last modification.</summary>
    ///<param name=""></param>
  }
, 'Items' : function() {
    ///<summary>Retrieves a FolderItems object that represents the collection of items in the folder.</summary>
    ///<returns type="FolderItems"/>
    return new FolderItems
  }
, 'MoveHere' : function() {
    ///<summary>Moves an item or items to this folder.</summary>
    ///<param name=""></param>
  }
, 'NewFolder' : function() {
    ///<summary>Creates a new folder.</summary>
    ///<param name=""></param>
  }
, 'ParseName' : function() {
    ///<summary>Creates and returns a FolderItem object that represents a specified item.</summary>
    ///<param name=""></param>
  }
, 'ParentFolder' : new Folder // Contains the parent Folder object.
, 'Title'        : ''         // Contains the title of the folder.
}

// FolderItem                                                     // http://msdn.microsoft.com/en-us/library/bb787810%28v=vs.85%29.aspx

function FolderItem() {}
FolderItem.prototype = {constructor:FolderItem
, 'InvokeVerb' : function(verb) {
    ///<summary>Executes a verb on the item.</summary>
    ///<param name="verb">
    ///A string that specifies the verb to be executed.
    ///It must be one of the values returned by the item's FolderItemVerb.Name property.
    ///If no verb is specified, the default verb will be invoked.
    ///Typically "open"
    ///</param>
  }
, 'Verbs' : function() {
    ///<summary>Retrieves the item's FolderItemVerbs object. This object is the collection of verbs that can be executed on the item.</summary>
    ///<returns type="FolderItemVerbs"/>
    return new FolderItemVerbs
  }
, 'GetFolder'    : new Folder // Contains the item's Folder object, if the item is a folder.
, 'GetLink'      : new Link   // Contains the item's ShellLinkObject object, if the item is a shortcut.
, 'IsBrowsable'  : false      // Indicates if the item can be hosted inside a browser or Windows Explorer frame.
, 'IsFileSystem' : false      // Indicates if the item is part of the file system.
, 'IsFolder'     : false      // Indicates if the item is a folder.
, 'IsLink'       : false      // Indicates whether the item is a shortcut.
, 'ModifyDate'   : ''         // Sets or gets the date and time that a file was last modified. ModifyDate can be used to retrieve the date and time that a folder was last modified, but cannot set it. "01/01/1900 6:05:00 PM"
, 'Name'         : ''         // Sets or gets the item's name.
, 'Parent'       : null       // Gets an object that represents the parent of the item.
, 'Path'         : ''         // Contains the item's full path and name.
, 'Size'         : 0          // Contains the item's size.
, 'Type'         : 0          // Contains a string representation of the item's type.
}

// FolderItems                                                    // http://msdn.microsoft.com/en-us/library/bb787800%28v=vs.85%29.aspx

function FolderItems() {}
FolderItems.prototype = {constructor:FolderItems
, 'Item' : function(index) {
    ///<summary>Retrieves the FolderItem object for a specified item in the collection.</summary>
    ///<param name="index">number, zero-based index of the item to retrieve. This value must be less than the value of the Count property.</param>
    ///<returns type="FolderItem"/>
    return new FolderItem
  }
, 'Count' : 0
}

// FolderItemVerb                                                 // http://msdn.microsoft.com/en-us/library/bb774172%28v=vs.85%29.aspx

function FolderItemVerb() {}
FolderItemVerb.prototype = {constructor:FolderItemVerb
, 'DoIt' : function() {
    ///<summary>Executes a verb on the FolderItem associated with the verb.</summary>
  }
, 'Name' : '' // Contains the verb's name.
}

// FolderItemVerbs                                                // http://msdn.microsoft.com/en-us/library/bb774160%28v=vs.85%29.aspx

function FolderItemVerbs() {}
FolderItemVerbs.prototype = {constructor:FolderItemVerbs
, 'Item' : function() {
    ///<summary>Retrieves the FolderItem object for a specified item in the collection.</summary>
    ///<param name="index">number, zero-based index of the item to retrieve. This value must be less than the value of the Count property.</param>
    ///<returns type="FolderItemVerb"/>
    return new FolderItemVerb
  }
, 'Count' : 0
}

// Link                                                           // http://msdn.microsoft.com/en-us/library/bb774004%28v=vs.85%29

function Link() {}
Link.prototype = {constructor:Link
, 'GetIconLocation' : function() {
    ///<summary>Gets the location of the icon assigned to the link.</summary>
  }
, 'Resolve' : function() {
    ///<summary>Looks for the target of a Shell link, even if the target has been moved or renamed.</summary>
  }
, 'Save' : function() {
    ///<summary>Saves all changes to the link.</summary>
  }
, 'SetIconLocation' : function() {
    ///<summary>Sets the location of the icon assigned to the link.</summary>
  }
, 'Arguments'        : '' // Contains a link's arguments.
, 'Description'      : '' // Gets or sets the description of the link.
, 'Hotkey'           : '' // Gets or sets the keyboard shortcut for the link.
, 'Path'             : '' // Gets or sets the path to the link object.
, 'ShowCommand'      : '' // Gets or sets the initial display state (sized, minimized, or maximized) of the link's command.
, 'WorkingDirectory' : '' // Gets or sets the working directory specified in the link.
}
}()
!function() { ActiveXObject['wscript.network'] = c
function c() {}
c.prototype = {constructor:c
, 'AddWindowsPrinterConnection' : function(path) {
    ///<summary>Adds a Windows-based printer connection to your computer system.&#10;Using this method is similar to using the Printer option on Control Panel to add a printer connection. Unlike the AddPrinterConnection method, this method allows you to create a printer connection without directing it to a specific port, such as LPT1. If the connection fails, an error is thrown.</summary>
    ///<param name="path">string, the path to the printer connection.&#10;"\\\\printserv\\DefaultPrinter"</param>
  }
, 'AddPrinterConnection' : function(local, remote, updateProfile, user, password) {
    ///<summary>Adds a remote MS-DOS-based printer connection to your computer system.</summary>
    ///<param name="local">String value indicating the local name to assign to the connected printer. &#10; "LPT1"</param>
    ///<param name="remote">String value indicating the name of the remote printer.&#10; "\\\\Server\\Print1"</param>
    ///<param name="updateProfile">Optional. Boolean value indicating whether the printer mapping is stored in the current user's profile. If updateProfile is supplied and is true, the mapping is stored in the user profile. The default value is false.</param>
    ///<param name="user">Optional. String value indicating the user name. If you are mapping a remote printer using the profile of someone other than current user, you can specify user and password.</param>
    ///<param name="password">Optional. String value indicating the user password. If you are mapping a remote printer using the profile of someone other than current user, you can specify user and password.</param>
  }
, 'EnumNetworkDrives' : function() {
    ///<summary>Returns the current network drive mapping information.</summary>
    ///<returns type="Collection"/>
    return new Collection
  }
, 'EnumPrinterConnections' : function() {
    ///<summary>Returns the current network printer mapping information.</summary>
    ///<returns type="Collection"/>
    return new Collection
  }
, 'MapNetworkDrive' : function(local, remote, updateProfile, user, password) {
    ///<summary>Adds a shared network drive to your computer system.</summary>
    ///<param name="local">String value indicating the name by which the mapped drive will be known locally.&#10;"s:"</param>
    ///<param name="remote">String value indicating the share's UNC name (\\xxx\yyy).&#10;"\\\\dns1\\dfs"</param>
    ///<param name="updateProfile">Optional. Boolean value indicating whether the mapping information is stored in the current user's profile.</param>
    ///<param name="user">Optional. String value indicating the user name. You must supply this argument if you are mapping a network drive using the credentials of someone other than the current user.</param>
    ///<param name="password">Optional. String value indicating the user password. You must supply this argument if you are mapping a network drive using the credentials of someone other than the current user.</param>
  }
, 'RemoveNetworkDrive' : function(name, force, updateProfile) {
    ///<summary>Removes a shared network drive from your computer system.</summary>
    ///<param name="name">String value indicating the name of the mapped drive you want to remove. The strName parameter can be either a local name or a remote name depending on how the drive is mapped. </param>
    ///<param name="force">Optional. Boolean value indicating whether to force the removal of the mapped drive. If bForce is supplied and its value is true, this method removes the connections whether the resource is used or not.</param>
    ///<param name="updateProfile">Optional. String value indicating whether to remove the mapping from the user's profile. If updateProfile is supplied and its value is true, this mapping is removed from the user profile. updateProfile is false by default.</param>
  }
, 'RemovePrinterConnection' : function(name, force, updateProfile) {
    ///<summary>Removes a shared network printer connection from your computer system.</summary>
    ///<param name="name">String value indicating the name that identifies the printer. It can be a UNC name (in the form \\xxx\yyy) or a local name (such as LPT1).</param>
    ///<param name="force">Optional. Boolean value indicating whether to force the removal of the mapped printer. If set to true (the default is false), the printer connection is removed whether or not a user is connected.</param>
    ///<param name="updateProfile">Optional. Boolean value. If set to true (the default is false), the change is saved in the user's profile.</param>
  }
, 'SetDefaultPrinter' : function(printerName) {
    ///<summary>Assigns a remote printer the role Default Printer.</summary>
    ///<param name="printerName">String value indicating the remote printer's UNC name.&#10;"\\\\research\\library1"</param>
  }
, 'ComputerName' : '' // Returns the name of the computer system.
, 'UserDomain'   : '' // Returns a user's domain name.
, 'UserName'     : '' // Returns the name of a user.
}

// Collection                                                     //

function Collection() {}
Collection.prototype = {constructor:Collection
, 'Count' : function() {
    ///<returns type="Number"/>
  }
, 'Item' : function(index) {
    ///<summary>
    ///&#10;This collection is an array that associates pairs of items — local names and their associated UNC names.
    ///&#10;Even-numbered items in the collection represent local names.
    ///&#10;Odd-numbered items represent the associated UNC share names.
    ///&#10;The first item in the collection is at index zero (0).
    ///</summary>
    ///<returns type="String"/>
  }
}

}()
!function() { ActiveXObject['wscript.shell'] = c
function c() {}
c.prototype = {constructor:c
, 'AppActivate' : function(title) {
    ///<summary>Activates an application window.</summary>
    ///<param name="title">Specifies which application to activate. This can be a string containing the title of the application (as it appears in the title bar) or the application's Process ID. </param>
    ///<returns type="Boolean"/>
  }
, 'CreateShortcut' : function(path) {
    ///<summary>Creates a new shortcut, or opens an existing shortcut.</summary>
    ///<param name="path">String value indicating the pathname of the shortcut to create.</param>
    return /\.url$/i.test(path)
      ? new WshURLShortcut
      : new WshShortcut
  }
, 'CurrentDirectory' : '' // The CurrentDirectory returns a string that contains the fully qualified path of the current working directory of the active process.
, 'Environment' : function(type) {
    ///<summary>Returns the WshEnvironment object (a collection of environment variables).</summary>
    ///<param name="type">SYSTEM, USER, VOLATILE, or PROCESS.</param>
    return function(name) {
      ///<param name="name">
      ///&#10;NUMBER_OF_PROCESSORS
      ///&#10;PROCESSOR_ARCHITECTURE
      ///&#10;PROCESSOR_IDENTIFIER
      ///&#10;PROCESSOR_LEVEL
      ///&#10;PROCESSOR_REVISION
      ///&#10;OS
      ///&#10;COMSPEC ; Executable file for the command prompt (cmd.exe)
      ///&#10;HOMEDRIVE ; Primary local drive (C)
      ///&#10;HOMEPATH ; Default directory for users (\users\default)
      ///&#10;PATH
      ///&#10;PATHEXT ; Extensions for executable files (.com, .exe, .bat, or .cmd)
      ///&#10;PROMPT  ; Command prompt (typically $P$G).
      ///&#10;SYSTEMDRIVE
      ///&#10;SYSTEMROOT
      ///&#10;WINDIR
      ///&#10;TEMP
      ///&#10;TMP
      ///</param>
    }
  }
, 'Exec' : function(command) {
    ///<summary>Runs an application in a child command-shell, providing access to the StdIn/StdOut/StdErr streams.</summary>
    ///<param name="command">String value indicating the command line used to run the script. The command line should appear exactly as it would if you typed it at the command prompt.</param>
    ///<returns type="WshScriptExec"/>
    return new WshScriptExec
  }
, 'ExpandEnvironmentStrings' : function(environmentVariable) {
    ///<summary>Returns an environment variable's expanded value.</summary>
    ///<param name="environmentVariable">String value indicating the name of the environment variable you want to expand. </param>
    ///<returns type="String"/>
  }
, 'LogEvent' : function(type, message, target) {
    ///<summary>Adds an event entry to a log file.</summary>
    ///<param name="type">
    ///&#10;0 SUCCESS
    ///&#10;1 ERROR
    ///&#10;2 WARNING
    ///&#10;4 INFORMATION
    ///&#10;8 AUDIT_SUCCESS
    ///&#10;16 AUDIT_FAILURE
    ///</param>
    ///<param name="message">String value containing the log entry text. </param>
    ///<param name="target">Optional. String value indicating the name of the computer system where the event log is stored (the default is the local computer system). Applies to Windows NT/2000 only. </param>
    ///<returns type="Boolean"/>
  }
, 'Popup' : function(text, wait, title, type) {
    ///<summary>Displays text in a pop-up message box.</summary>
    ///<param name="text">String value containing the text you want to appear in the pop-up message box.</param>
    ///<param name="wait">Optional. Numeric value indicating the maximum length of time (in seconds) you want the pop-up message box displayed. </param>
    ///<param name="title">Optional. String value containing the text you want to appear as the title of the pop-up message box. </param>
    ///<param name="type">Optional. Numeric value indicating the type of buttons and icons you want in the pop-up message box. These determine how the message box is used.
    ///&#10;Button Types
    ///&#10; 0 Show OK button. 
    ///&#10; 1 Show OK and Cancel buttons. 
    ///&#10; 2 Show Abort, Retry, and Ignore buttons. 
    ///&#10; 3 Show Yes, No, and Cancel buttons. 
    ///&#10; 4 Show Yes and No buttons. 
    ///&#10; 5 Show Retry and Cancel buttons.
    ///&#10;Icon Types
    ///&#10; 16 Show "Stop Mark" icon. 
    ///&#10; 32 Show "Question Mark" icon. 
    ///&#10; 48 Show "Exclamation Mark" icon. 
    ///&#10; 64 Show "Information Mark" icon. 
    ///</param>
    ///<returns type="Number">Integer value indicating the number of the button the user clicked to dismiss the message box. This is the value returned by the Popup method.
    ///&#10; 1 OK button 
    ///&#10; 2 Cancel button 
    ///&#10; 3 Abort button 
    ///&#10; 4 Retry button 
    ///&#10; 5 Ignore button 
    ///&#10; 6 Yes button 
    ///&#10; 7 No button 
    ///</returns>
  }
, 'RegDelete' : function(name) {
    ///<summary>Deletes a key or one of its values from the registry.</summary>
    ///<param name="name">String value indicating the name of the registry key or key value you want to delete.
    ///&#10;Specify a key-name by ending "name" with a final backslash; leave it off to specify a value-name.
    ///&#10;Fully qualified key-names and value-names are prefixed with a root key.
    ///&#10;You may use abbreviated versions of root key names with the RegDelete method.
    ///&#10;The five possible root keys you can use are listed in the following table.
    ///&#10;HKCU = HKEY_CURRENT_USER
    ///&#10;HKLM = HKEY_LOCAL_MACHINE
    ///&#10;HKCR = HKEY_CLASSES_ROOT
    ///&#10;HKEY_USERS = (same)
    ///&#10;HKEY_CURRENT_CONFIG = (same)
    ///</param>
  }
, 'RegRead' : function(name) {
    ///<summary>Returns the value of a key or value-name from the registry.</summary>
    ///<param name="name">String value indicating the key or value-name whose value you want.
    ///&#10;HKCU = HKEY_CURRENT_USER
    ///&#10;HKLM = HKEY_LOCAL_MACHINE
    ///&#10;HKCR = HKEY_CLASSES_ROOT
    ///&#10;HKEY_USERS = (same)
    ///&#10;HKEY_CURRENT_CONFIG = (same)
    ///</param>
    ///<returns type="Variant">
    ///&#10; REG_SZ          A string A string 
    ///&#10; REG_DWORD       A number An integer 
    ///&#10; REG_BINARY      A binary value A VBArray of integers 
    ///&#10; REG_EXPAND_SZ   An expandable string (e.g., "%windir%\\calc.exe") A string 
    ///&#10; REG_MULTI_SZ    An array of strings A VBArray of strings 
    ///</returns>
  }
, 'RegWrite' : function(name, value, type) {
    ///<summary>Creates a new key, adds another value-name to an existing key (and assigns it a value), or changes the value of an existing value-name.</summary>
    ///<param name="name">String value indicating the key-name, value-name, or value you want to create, add, or change. 
    ///&#10;HKCU = HKEY_CURRENT_USER
    ///&#10;HKLM = HKEY_LOCAL_MACHINE
    ///&#10;HKCR = HKEY_CLASSES_ROOT
    ///&#10;HKEY_USERS = (same)
    ///&#10;HKEY_CURRENT_CONFIG = (same)
    ///</param>
    ///<param name="value">The name of the new key you want to create, the name of the value you want to add to an existing key, or the new value you want to assign to an existing value-name.</param>
    ///<param name="type">Optional. String value indicating the value's data type.
    ///&#10;REG_SZ = string 
    ///&#10;REG_BINARY = number
    ///&#10;REG_DWORD  = number 
    ///&#10;REG_EXPAND_SZ = string
    ///&#10;If undefined:
    ///&#10; String = REG_SZ | REG_EXPAND_SZ 
    ///&#10; Integer = REG_DWORD | REG_BINARY 
    ///</param>
  }
, 'Run' : function(command, windowStyle, wait) {
    ///<summary>Runs a program in a new process.</summary>
    ///<param name="command">string, String value indicating the command line you want to run. You must include any parameters you want to pass to the executable file. 
    ///&#10; "notepad "+ WScript.ScriptFullName, 1, true
    ///&#10; "cmd /K CD C:\ & Dir"
    ///</param>
    ///<param name="windowStyle">number, Optional. Integer value indicating the appearance of the program's window. Note that not all programs make use of this information. 
    ///&#10; 0 Hides the window and activates another window. 
    ///&#10; 1 Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when displaying the window for the first time. 
    ///&#10; 2 Activates the window and displays it as a minimized window.  
    ///&#10; 3 Activates the window and displays it as a maximized window.  
    ///&#10; 4 Displays a window in its most recent size and position. The active window remains active. 
    ///&#10; 5  Activates the window and displays it in its current size and position. 
    ///&#10; 6 Minimizes the specified window and activates the next top-level window in the Z order. 
    ///&#10; 7 Displays the window as a minimized window. The active window remains active. 
    ///&#10; 8 Displays the window in its current state. The active window remains active. 
    ///&#10; 9 Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window. 
    ///&#10; 10 Sets the show-state based on the state of the program that started the application. 
    ///</param>
    ///<param name="wait">boolean, Optional. Boolean value indicating whether the script should wait for the program to finish executing before continuing to the next statement in your script. If set to true, script execution halts until the program finishes, and Run returns any error code returned by the program. If set to false (the default), the Run method returns immediately after starting the program, automatically returning 0 (not to be interpreted as an error code). </param>
    ///<returns type="Number"/>
  }
, 'SendKeys' : function(string) {
    ///<summary>Sends one or more keystrokes to the active window (as if typed on the keyboard).</summary>
    ///<param name="string">String value indicating the keystroke(s) you want to send.
    ///&#10;BACKSPACE {BACKSPACE}, {BS}, or {BKSP} 
    ///&#10;BREAK {BREAK} 
    ///&#10;CAPS LOCK {CAPSLOCK} 
    ///&#10;DEL or DELETE {DELETE} or {DEL} 
    ///&#10;DOWN ARROW {DOWN} 
    ///&#10;END {END} 
    ///&#10;ENTER {ENTER} or ~ 
    ///&#10;ESC {ESC} 
    ///&#10;HELP {HELP} 
    ///&#10;HOME {HOME} 
    ///&#10;INS or INSERT {INSERT} or {INS} 
    ///&#10;LEFT ARROW {LEFT} 
    ///&#10;NUM LOCK {NUMLOCK} 
    ///&#10;PAGE DOWN {PGDN} 
    ///&#10;PAGE UP {PGUP} 
    ///&#10;PRINT SCREEN {PRTSC} 
    ///&#10;RIGHT ARROW {RIGHT} 
    ///&#10;SCROLL LOCK {SCROLLLOCK} 
    ///&#10;TAB {TAB} 
    ///&#10;UP ARROW {UP} 
    ///&#10;F1 to F16 {F1} to {F16}
    ///&#10;SHIFT + 
    ///&#10;CTRL ^ 
    ///&#10;ALT % 
    ///</param>
  }
, 'SpecialFolders' : function(specialFolder) {
    ///<summary>Returns a SpecialFolders object (a collection of special folders).</summary>
    ///<param name="specialFolder">
    ///&#10;AllUsersDesktop
    ///&#10;AllUsersStartMenu
    ///&#10;AllUsersPrograms
    ///&#10;AllUsersStartup
    ///&#10;Desktop
    ///&#10;Favorites
    ///&#10;Fonts
    ///&#10;MyDocuments
    ///&#10;NetHood
    ///&#10;PrintHood
    ///&#10;Programs
    ///&#10;Recent
    ///&#10;SendTo
    ///&#10;StartMenu
    ///&#10;Startup
    ///&#10;Templates
    ///</param>
    ///<returns type="String"/>
  }
}

// TextStream                                                     //

function TextStream() {}
TextStream.prototype = {constructor:TextStream
, 'Close' : function() {
    ///<summary>Closes an open TextStream file.</summary>
  }
, 'Read' : function(characters) {
    ///<summary>Reads a specified number of characters from a TextStream file and returns the resulting string.</summary>
    ///<param name="characters">Required. Number of characters you want to read from the file.</param>
    ///<returns type="String"/>
  }
, 'ReadAll' : function() {
    ///<summary>Reads an entire TextStream file and returns the resulting string.</summary>
    ///<returns type="String"/>
  }
, 'ReadLine' : function() {
    ///<summary>Reads an entire line (up to, but not including, the newline character) from a TextStream file and returns the resulting string.</summary>
    ///<returns type="String"/>
  }
, 'Skip' : function(characters) {
    ///<summary>Skips a specified number of characters when reading a TextStream file.</summary>
    ///<param name="characters">Required. Number of characters to skip when reading a file.</param>
  }
, 'SkipLine' : function() {
    ///<summary>Skips the next line when reading a TextStream file.</summary>
  }
, 'Write' : function(string) {
    ///<summary>Writes a specified string to a TextStream file.</summary>
    ///<param name="string">Required. The text you want to write to the file.</param>
  }
, 'WriteBlankLines' : function(lines) {
    ///<summary>Writes a specified number of newline characters to a TextStream file.</summary>
    ///<param name="lines">Required. Number of newline characters you want to write to the file.</param>
  }
, 'WriteLine':function(string) {
    ///<summary>Writes a specified string and newline character to a TextStream file.</summary>
    ///<param name="string">Optional. The text you want to write to the file. If omitted, a newline character is written to the file.</param>
  }
, 'AtEndOfLine'   : true // Returns true if the file pointer is positioned immediately before the end-of-line marker in a TextStream file; false if it is not. Read-only.
, 'AtEndOfStream' : true // Returns true if the file pointer is at the end of a TextStream file; false if it is not. Read-only.
, 'Column'        : 0    // Read-only property that returns the column number of the current character position in a TextStream file.
, 'Line'          : 0    // Read-only property that returns the current line number in a TextStream file.
}

// WshScriptExec                                                  //

function WshScriptExec() {}
WshScriptExec.prototype = {constructor:WshScriptExec
, 'Terminate' : function() {
    ///<summary>Instructs the script engine to end the process started by the Exec method.</summary>
  }
, 'Status' : 0 // WshRunning ( = 0) The job is still running. WshFinished ( = 1) The job has completed.
, 'StdOut' : new TextStream
, 'StdIn'  : new TextStream
, 'StdErr' : new TextStream
}

// WshShortcut                                                    //

function WshShortcut() {}
WshShortcut.prototype = {constructor:WshShortcut
, 'Save' : function() {
    ///<summary>Saves a shortcut object to disk.</summary>
  }
, 'Arguments'        : '' // 
, 'Description'      : '' // The Description property contains a string value describing a shortcut.
, 'FullName'         : '' // The FullName property contains a read-only string value indicating the fully qualified path to the shortcut's target.
, 'Hotkey'           : '' // Assigns a key-combination to a shortcut. "ALT+ , CTRL+ , SHIFT+ , EXT+"
, 'IconLocation'     : '' // The string should contain a fully qualified path and an index associated with the icon. "notepad.exe, 0"
, 'TargetPath'       : '' // This property is for the shortcut's target path only. Any arguments to the shortcut must be placed in the Argument's property.
, 'WindowStyle'      : '' // Sets the window style for the program being run.
                          // 1 Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. 
                          // 3 Activates the window and displays it as a maximized window.  
                          // 7 Minimizes the window and activates the next top-level window. 
, 'WorkingDirectory' : '' // Directory in which the shortcut starts. SpecialFolders("Desktop")
}

// WshURLShortcut                                                 //

function WshURLShortcut() {}
WshURLShortcut.prototype = {constructor:WshURLShortcut
, 'Save' : function() {
    ///<summary>Saves a shortcut object to disk.</summary>
  }
, 'TargetPath' : ''
}

}()
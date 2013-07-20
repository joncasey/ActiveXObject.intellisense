///<reference path="ActiveXObject.js">
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
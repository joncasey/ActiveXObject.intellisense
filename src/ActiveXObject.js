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
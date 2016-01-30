
/*
 * ActiveXObject Intellisense file for Visual Studio
 * Copyright Jon Casey <joncasey@gmail.com>
 * Released under the MIT license
 */

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

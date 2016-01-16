
var outputFile = 'ActiveXObject.intellisense.js'
var sourceFiles = [
  'ActiveXObject',
  'ADODB.Stream',
  'MSXML',
  'Scripting.FileSystemObject',
  'Shell.Application',
  'WScript.Network',
  'WScript.Shell'
]

var fs = new ActiveXObject('Scripting.FileSystemObject')

for (var i = 0; i < sourceFiles.length; i++) {
  var textStream = fs.OpenTextFile('src/' + sourceFiles[i] + '.js')
  var text = textStream.ReadAll()
  textStream.Close()
  sourceFiles[i] = text.replace(/^\/{3}.+\n/gm, '')
}

with (fs.CreateTextFile(outputFile, true)) {
  Write(sourceFiles.join('\r\n'))
  Close()
}

WScript.Echo('Saved "' + outputFile + '"')










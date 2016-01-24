
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
  sourceFiles[i] = read('src/' + sourceFiles[i] + '.js').replace(/^\/{3}.+\n/gm, '')
}

sourceFiles.unshift(
  read('doc/Copyright.txt').replace(/YYYY/, new Date().getFullYear())
)

with (fs.CreateTextFile(outputFile, true)) {
  Write(sourceFiles.join('\r\n'))
  Close()
}

WScript.Echo('Saved "' + outputFile + '"')



function read(file) {
  var stream = fs.OpenTextFile(file)
  var text = stream.ReadAll()
  stream.Close()
  return text
}






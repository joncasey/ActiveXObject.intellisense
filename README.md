# ActiveXObject.intellisense

This project provides IntelliSense for ActiveObject(s) in Visual Studio.
I got tired of waiting for Microsoft to add intellisense for them, so I wrote my own.

This works in Visual Studio & Visual Studio Express from 2008-2015.
This will not work for VSCode. It does not support this type of IntelliSense markup.

## Installation

I don't have time to make a fancy installer.
But, the process is really pretty easy.

1. Copy the [dist/ActiveXObject.intellisense.js](https://raw.githubusercontent.com/joncasey/ActiveXObject.intellisense/master/dist/ActiveXObject.intellisense.js) to your {{Visual Studio}}/JavaScript/References folder. (or anywhere you want)
2. Open Tools / Options
3. Goto Text Editor / JavaScript / IntelliSense / References
4. Toggle Reference Group to "Implicit (Web)"
5. Click the "…" button to "Add a reference to current group"
6. Select the location for the ActiveXObject.intellisense.js you copied in step #1.

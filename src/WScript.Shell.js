///<reference path="ActiveXObject.js">
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
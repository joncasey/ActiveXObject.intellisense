/// <reference path="ActiveXObject.js">
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
      /// "cmd /K CD C:\ & Dir"                         &#10;
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

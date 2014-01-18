var WScript = WScript || function() {

  return new WScript

  function WScript() {
    /// <field name="Arguments" type="WshArguments"></field>
    /// <field name="FullName" type="String">Returns the fully qualified path of the host executable (CScript.exe or WScript.exe).</field>
    /// <field name="Interactive" type="String">Sets the script mode this. or identifies the script mode.</field>
    /// <field name="Name" type="String">Returns the name of the WScript object (the host executable file).</field>
    /// <field name="Path" type="String">Returns the name of the directory containing the host executable (CScript.exe or WScript.exe).</field>
    /// <field name="ScriptFullName" type="String">Returns the full path of the currently running script.</field>
    /// <field name="ScriptName" type="String">Returns the file name of the currently running script.</field>
    /// <field name="StdErr" type="String">Exposes the write-only error output stream for the current script.</field>
    /// <field name="StdIn" type="String">Exposes the read-only input stream for the current script.</field>
    /// <field name="StdOut" type="String">Exposes the write-only output stream for the current script.</field>
    /// <field name="Version" type="String">Returns the version of Windows Script Host.</field>
    this.Arguments = new WshArguments
    this.FullName = ''
    this.Interactive = ''
    this.Name = ''
    this.Path = ''
    this.ScriptFullName = ''
    this.ScriptName = ''
    this.StdErr = ''
    this.StdIn = ''
    this.StdOut = ''
    this.Version = ''
    this.CreateObject = function(progID, prefix) {
      /// <summary>Creates a COM object.</summary>
      /// <param name="progID">String value indicating the programmatic identifier (ProgID) of the object you want to create.</param>
      /// <param name="prefix" optional="true">Optional. String value indicating the function prefix.</param>
      /*
        Objects created with the CreateObject method using the strPrefix argument are connected objects.
        These are useful when you want to sync an object's events.
        The object's outgoing interface is connected to the script file after the object is created.
        Event functions are a combination of this prefix and the event name.
        If you create an object and do not supply the strPrefix argument, you can still sync events on the object by using the ConnectObject method.
        When the object fires an event, WSH calls a subroutine with strPrefix attached to the beginning of the event name.
        For example, if strPrefix is MYOBJ and the object fires an event named OnBegin, Windows Script Host calls the MYOBJ_OnBegin subroutine located in the script.
        The CreateObject method returns a pointer to the object's IDispatch interface.
      */
    }
    this.ConnectObject = function(object, prefix) {
      /// <summary>Connects the object's event sources to functions with a given prefix.</summary>
      /// <param name="object">Required. String value indicating the name of the object you want to connect.</param>
      /// <param name="prefix">Required. String value indicating the function prefix.</param>
      /*
        Connected objects are useful when you want to sync an object's events.
        The ConnectObject method connects the object's outgoing interface to the script file after creating the object.
        Event functions are a combination of this prefix and the event name.
      */
    }
    this.DisconnectObject = function(object) {
      /// <summary>Disconnects a connected object's event sources.</summary>
      /// <param name="object">String value indicating the name of the object to disconnect.</param>
      /*
        Once an object has been "disconnected," WSH will not respond to its events.
        The object is still capable of firing events, though.
        Note that the DisconnectObject method does nothing if the specified object is not already connected.
      */
    }
    this.Echo = function(arg1, arg2, argN) {
      /// <summary>Outputs text to either a message box or the command console window.</summary>
      /// <param name="arg1" optional="true">Optional. String value indicating the list of items to be displayed.</param>
    }
    this.GetObject = function(path, progID, prefix) {
      /// <summary>Retrieves an existing object with the specified ProgID, or creates a new one from a file.</summary>
      /// <param name="path">The fully qualified path name of the file that contains the object persisted to disk.</param>
      /// <param name="progID" optional="true">Optional. The object's program identifier (ProgID).</param>
      /// <param name="prefix" optional="true">Optional. Used when you want to sync the object's events. If you supply the strPrefix argument, WSH connects the object's outgoing interface to the script file after creating the object.</param>
   }
    this.Quit = function(errorCode) {
      /// <summary>Forces script execution to stop at any time.</summary>
      /// <param name="errorCode" optional="true">Optional. Integer value returned as the process's exit code. If you do not include the intErrorCode parameter, no value is returned.</param>
    }
    this.Sleep = function(time) {
      /// <summary>Suspends script execution for a specified length of time, then continues execution.</summary>
      /// <param name="time">Integer value indicating the interval (in milliseconds) you want the script process to be inactive.</param>
    }
  }

  function WshArguments() {
    /// <summary>Provides access to the entire collection of command-line parameters — in the order in which they were originally entered.</summary>
    /// <field name="Length" type="Number">Returns the number of command-line parameters belonging to a script (the number of items in an argument's collection).</field>
    this.Count = function() {}
    this.Item = function(key) {
      /// <summary>Provides access to the items in the WshNamed object.</summary>
      /// <param name="key">The name or index of the item you want to retrieve.</param>
      /// <returns type="String"/>
    }
    this.Length = 0
    this.Named = new WshNamed
    this.Unnnamed = new WshUnnamed
    this.ShowUsage = function() {
      /// <summary>Makes a script self-documenting by displaying information about how it should be used.</summary>
    }
  }

  function WshNamed() {
    /// <field name="Length" type="Number">Returns the number of command-line parameters belonging to a script (the number of items in an argument's collection).</field>
    this.Count = function() {}
    this.Exists = function(key) {
      /// <summary>Indicates whether a specific key value exists in the WshNamed object.</summary>
      /// <param name="key">String value indicating an argument of the WshNamed object.</param>
      /// <returns type="Boolean"/>
    }
    this.Item = function(key) {
      /// <summary>Provides access to the items in the WshNamed object.</summary>
      /// <param name="key">The name of the item you want to retrieve.</param>
      /// <returns type="String"/>
    }
    this.Length = 0
  }

  function WshUnnamed() {
    /// <field name="Length" type="Number">Returns the number of command-line parameters belonging to a script (the number of items in an argument's collection).</field>
    this.Count = function() {}
    this.Item = function(index) {
      /// <summary>Provides access to the items in the WshUnnamed object.</summary>
      /// <param name="index">The number of the item you want to retrieve. </param>
      /// <returns type="String"/>
    }
    this.Length = 0
  }

}();

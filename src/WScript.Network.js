/// <reference path="ActiveXObject.js">
+function() {

  ActiveXObject['wscript.network'] = WScriptNetwork

  function WScriptNetwork() {
    /// <field name="ComputerName" type="String">Returns the name of the computer system.</field>
    /// <field name="UserDomain" type="String">Returns a user's domain name.</field>
    /// <field name="UserName" type="String">Returns the name of a user.</field>
    this.ComputerName = ''
    this.UserDomain = ''
    this.UserName = ''
    this.AddWindowsPrinterConnection = function(path) {
      /// <summary>Adds a Windows-based printer connection to your computer system.&#10;Using this method is similar to using the Printer option on Control Panel to add a printer connection. Unlike the AddPrinterConnection method, this method allows you to create a printer connection without directing it to a specific port, such as LPT1. If the connection fails, an error is thrown.</summary>
      /// <param name="path">string, the path to the printer connection.&#10;"\\\\printserv\\DefaultPrinter"</param>
    }
    this.AddPrinterConnection = function(local, remote, updateProfile, user, password) {
      /// <summary>Adds a remote MS-DOS-based printer connection to your computer system.</summary>
      /// <param name="local">String value indicating the local name to assign to the connected printer. &#10; "LPT1"</param>
      /// <param name="remote">String value indicating the name of the remote printer.&#10; "\\\\Server\\Print1"</param>
      /// <param name="updateProfile" optional="true">Optional. Boolean value indicating whether the printer mapping is stored in the current user's profile. If updateProfile is supplied and is true, the mapping is stored in the user profile. The default value is false.</param>
      /// <param name="user" optional="true">Optional. String value indicating the user name. If you are mapping a remote printer using the profile of someone other than current user, you can specify user and password.</param>
      /// <param name="password" optional="true">Optional. String value indicating the user password. If you are mapping a remote printer using the profile of someone other than current user, you can specify user and password.</param>
    }
    this.EnumNetworkDrives = function() {
      /// <summary>Returns the current network drive mapping information.</summary>
      /// <returns type="Collection"/>
      return new Collection
    }
    this.EnumPrinterConnections = function() {
      /// <summary>Returns the current network printer mapping information.</summary>
      /// <returns type="Collection"/>
      return new Collection
    }
    this.MapNetworkDrive = function(local, remote, updateProfile, user, password) {
      /// <summary>Adds a shared network drive to your computer system.</summary>
      /// <param name="local">String value indicating the name by which the mapped drive will be known locally.&#10;"s:"</param>
      /// <param name="remote">String value indicating the share's UNC name (\\xxx\yyy).&#10;"\\\\dns1\\dfs"</param>
      /// <param name="updateProfile" optional="true">Optional. Boolean value indicating whether the mapping information is stored in the current user's profile.</param>
      /// <param name="user" optional="true">Optional. String value indicating the user name. You must supply this argument if you are mapping a network drive using the credentials of someone other than the current user.</param>
      /// <param name="password" optional="true">Optional. String value indicating the user password. You must supply this argument if you are mapping a network drive using the credentials of someone other than the current user.</param>
    }
    this.RemoveNetworkDrive = function(name, force, updateProfile) {
      /// <summary>Removes a shared network drive from your computer system.</summary>
      /// <param name="name">String value indicating the name of the mapped drive you want to remove. The strName parameter can be either a local name or a remote name depending on how the drive is mapped. </param>
      /// <param name="force" optional="true">Optional. Boolean value indicating whether to force the removal of the mapped drive. If bForce is supplied and its value is true, this method removes the connections whether the resource is used or not.</param>
      /// <param name="updateProfile" optional="true">Optional. String value indicating whether to remove the mapping from the user's profile. If updateProfile is supplied and its value is true, this mapping is removed from the user profile. updateProfile is false by default.</param>
    }
    this.RemovePrinterConnection = function(name, force, updateProfile) {
      /// <summary>Removes a shared network printer connection from your computer system.</summary>
      /// <param name="name">String value indicating the name that identifies the printer. It can be a UNC name (in the form \\xxx\yyy) or a local name (such as LPT1).</param>
      /// <param name="force" optional="true">Optional. Boolean value indicating whether to force the removal of the mapped printer. If set to true (the default is false), the printer connection is removed whether or not a user is connected.</param>
      /// <param name="updateProfile" optional="true">Optional. Boolean value. If set to true (the default is false), the change is saved in the user's profile.</param>
    }
    this.SetDefaultPrinter = function(printerName) {
      /// <summary>Assigns a remote printer the role Default Printer.</summary>
      /// <param name="printerName">String value indicating the remote printer's UNC name.&#10;"\\\\research\\library1"</param>
    }
  }

  function Collection() {
    this.Count = function() {
      /// <returns type="Number"/>
    }
    this.Item = function(index) {
      /// <summary>
      /// This collection is an array that associates pairs of items — local names and their associated UNC names. &#10;
      /// Even-numbered items in the collection represent local names.                                             &#10;
      /// Odd-numbered items represent the associated UNC share names.                                             &#10;
      /// The first item in the collection is at index zero (0).                                                   &#10;
      /// </summary>
      /// <returns type="String"/>
    }
  }

}();

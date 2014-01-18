/// <reference path="ActiveXObject.js">
+function() {

  ActiveXObject['shell.application'] = ShellApplication

  function ShellApplication() {
    this.AddToRecent = function(file, category) {
      /// <summary>Adds a file to the most recently used (MRU) list.</summary>
      /// <param name="file">string, path of the file to add to the list of recently used documents&#10;null, to clear the recent documents folder</param>
      /// <param name="category">string, name of the category in which ot place the file (optional)</param>
    }
    this.BrowseForFolder = function(hwnd, title, options, rootFolder) {
      /// <summary>Creates a dialog box that enables the user to select a folder and then returns the selected folder's Folder object.</summary>
      /// <param name="hwnd">number, handle of the parent window of the dialog box, or zero.</param>
      /// <param name="title">string, represents the title displayed inside the Browse dialog box.</param>
      /// <param name="options">number, zero or a combination of the ulFlags in BROWSEINFO</param>
      /// <param name="rootFolder">user can not browse higher than this folder &#10;String, specifying a path &#10;Number, one of the ShellSpecialFolderConstants (ssf) &#10;Undefined, will use DESKTOP</param>
      /// <returns type="Folder"/>
      return new Folder
    }
    this.CanStartStopService = function() {
      /// <summary>Determines if the current user can start and stop the named service.</summary>
      /// <param name=""></param>
    }
    this.CascadeWindows = function() {
      /// <summary>Cascades all of the windows on the desktop. This method has the same effect as right-clicking the taskbar and selecting Cascade Windows.</summary>
      /// <param name=""></param>
    }
    this.ControlPanelItem = function() {
      /// <summary>Runs the specified Control Panel (*.cpl) application. If the application is already open, it will activate the running instance.</summary>
      /// <param name=""></param>
    }
    this.EjectPC = function() {
      /// <summary>Ejects the computer from its docking station. This is the same as clicking the Start menu and selecting Eject PC, if your computer supports this command.</summary>
      /// <param name=""></param>
    }
    this.Explore = function() {
      ///<summary>Opens a specified folder in a Windows Explorer window.</summary>
      ///<param name=""></param>
    }
    this.ExplorerPolicy = function() {
      /// <summary>Gets the value for a specified Internet Explorer policy.</summary>
      /// <param name=""></param>
    }
    this.FileRun = function() {
      /// <summary>Displays the Run dialog to the user. This method has the same effect as clicking the Start menu and selecting Run.</summary>
      /// <param name=""></param>
    }
    this.FindComputer = function() {
      /// <summary>Displays the Search Results: Computers dialog box. The dialog box shows the result of the search for a specified computer.</summary>
      /// <param name=""></param>
    }
    this.FindFiles = function() {
      /// <summary>Displays the Find: All Files dialog box. This is the same as clicking the Start menu and then selecting Search (or its equivalent under systems earlier than Windows XP.</summary>
      /// <param name=""></param>
    }
    this.FindPrinter = function() {
      /// <summary>Displays the Find Printer dialog box.</summary>
      /// <param name=""></param>
    }
    this.GetSetting = function() {
      /// <summary>Retrieves a global Shell setting.</summary>
      /// <param name=""></param>
    }
    this.GetSystemInformation = function() {
      /// <summary>Retrieves system information.</summary>
      /// <param name=""></param>
    }
    this.Help = function() {
      /// <summary>Displays the Windows Help and Support Center. This method has the same effect as clicking the Start menu and selecting Help and Support.</summary>
      /// <param name=""></param>
    }
    this.IsRestricted = function() {
      /// <summary>Retrieves a group's restriction setting from the registry.</summary>
      /// <param name=""></param>
    }
    this.IsServiceRunning = function() {
      /// <summary>Returns a value that indicates whether a particular service is running.</summary>
      /// <param name=""></param>
    }
    this.MinimizeAll = function() {
      /// <summary>Minimizes all of the windows on the desktop. This method has the same effect as right-clicking the taskbar and selecting Minimize All Windows on older systems or clicking the Show Desktop icon in the Quick Launch area of the taskbar in Windows 2000 or Windows XP.</summary>
      /// <param name=""></param>
    }
    this.NameSpace = function(dir) {
      /// <summary>Creates and returns a Folder object for the specified folder.</summary>
      /// <param name="dir">String, path of the folder &#10;Number, a value from the ShellSpecialFolderConstants (ssf)</param>
      return new Folder
    }
    this.Open = function() {
      /// <summary>Opens the specified folder.</summary>
      /// <param name=""></param>
    }
    this.RefreshMenu = function() {
      /// <summary>Refreshes the contents of the Start menu. Used only with systems preceding Windows XP.</summary>
      /// <param name=""></param>
    }
    this.ServiceStart = function() {
      /// <summary>Starts a named service.</summary>
      /// <param name=""></param>
    }
    this.ServiceStop = function() {
      /// <summary>Stops a named service.</summary>
      /// <param name=""></param>
    }
    this.SetTime = function() {
      /// <summary>Displays the Date and Time Properties dialog box. This method has the same effect as right-clicking the clock in the taskbar status area and selecting Adjust Date/Time.</summary>
      /// <param name=""></param>
    }
    this.ShellExecute = function() {
      /// <summary>Performs a specified operation on a specified file.</summary>
      /// <param name=""></param>
    }
    this.ShowBrowserBar = function() {
      /// <summary>Displays a browser bar.</summary>
      /// <param name=""></param>
    }
    this.ShutdownWindows = function() {
      /// <summary>Displays the Shut Down Windows dialog box. This is the same as clicking the Start menu and selecting Shut Down.</summary>
      /// <param name=""></param>
    }
    this.Suspend = function() {
      /// <summary></summary>
      /// <param name=""></param>
    }
    this.TileHorizontally = function() {
      /// <summary>Tiles all of the windows on the desktop horizontally. This method has the same effect as right-clicking the taskbar and selecting Tile Windows Horizontally.</summary>
      /// <param name=""></param>
    }
    this.TileVertically = function() {
      /// <summary>Tiles all of the windows on the desktop vertically. This method has the same effect as right-clicking the taskbar and selecting Tile Windows Vertically.</summary>
      /// <param name=""></param>
    }
    this.TrayProperties = function() {
      /// <summary>Displays the Taskbar and Start Menu Properties dialog box. This method has the same effect as right-clicking the taskbar and selecting Properties.</summary>
      /// <param name=""></param>
    }
    this.UndoMinimizeALL = function() {
      /// <summary>Restores all desktop windows to the same state they were in before the last MinimizeAll command. This method has the same effect as right-clicking the taskbar and selecting Undo Minimize All Windows on older systems or a second clicking of the Show Desktop icon in the Quick Launch area of the taskbar in Windows 2000 or Windows XP.</summary>
      /// <param name=""></param>
    }
    this.Windows = function() {
      /// <summary>Creates and returns a ShellWindows object. This object represents a collection of all of the open windows that belong to the Shell.</summary>
      /// <param name=""></param>
    }
    this.WindowsSecurity = function() {
      /// <summary>Displays the Windows Security dialog box.</summary>
      /// <param name=""></param>
    }
    this.WindowSwitcher = function() {
      /// <summary>Displays your open windows in a 3D stack that you can flip through.</summary>
      /// <param name=""></param>
    }

    this.ulFlags = {
      // Only return file system directories. If the user selects folders that are not part of the file system, the OK button is grayed. Note  The OK button remains enabled for "\\server" items, as well as "\\server\share" and directory items. However, if the user selects a "\\server" item, passing the PIDL returned by SHBrowseForFolder to SHGetPathFromIDList fails.
      BIF_RETURNONLYFSDIRS    :     1,
      // Do not include network folders below the domain level in the dialog box's tree view control.
      BIF_DONTGOBELOWDOMAIN   :     2,
      // Include a status area in the dialog box. The callback function can set the status text by sending messages to the dialog box. This flag is not supported when BIF_NEWDIALOGSTYLE is specified.
      BIF_STATUSTEXT          :     4,
      // Only return file system ancestors. An ancestor is a subfolder that is beneath the root folder in the namespace hierarchy. If the user selects an ancestor of the root folder that is not part of the file system, the OK button is grayed.
      BIF_RETURNFSANCESTORS   :     8,
      // Version 4.71. Include an edit control in the browse dialog box that allows the user to type the name of an item.
      BIF_EDITBOX             :    10,
      // Version 4.71. If the user types an invalid name into the edit box, the browse dialog box calls the application's BrowseCallbackProc with the BFFM_VALIDATEFAILED message. This flag is ignored if BIF_EDITBOX is not specified.
      BIF_VALIDATE            :    20,
      // Version 5.0. Use the new user interface. Setting this flag provides the user with a larger dialog box that can be resized. The dialog box has several new capabilities, including: drag-and-drop capability within the dialog box, reordering, shortcut menus, new folders, delete, and other shortcut menu commands. Note  If COM is initialized through CoInitializeEx with the COINIT_MULTITHREADED flag set, SHBrowseForFolder fails if BIF_NEWDIALOGSTYLE is passed.
      BIF_NEWDIALOGSTYLE      :    40,
      // Version 5.0. The browse dialog box can display URLs. The BIF_USENEWUI and BIF_BROWSEINCLUDEFILES flags must also be set. If any of these three flags are not set, the browser dialog box rejects URLs. Even when these flags are set, the browse dialog box displays URLs only if the folder that contains the selected item supports URLs. When the folder's IShellFolder::GetAttributesOf method is called to request the selected item's attributes, the folder must set the SFGAO_FOLDER attribute flag. Otherwise, the browse dialog box will not display the URL.
      BIF_BROWSEINCLUDEURLS   :    80,
      // Version 5.0. Use the new user interface, including an edit box. This flag is equivalent to BIF_EDITBOX | BIF_NEWDIALOGSTYLE. Note  If COM is initialized through CoInitializeEx with the COINIT_MULTITHREADED flag set, SHBrowseForFolder fails if BIF_USENEWUI is passed.
      BIF_USENEWUI            :    50,
      // Version 6.0. When combined with BIF_NEWDIALOGSTYLE, adds a usage hint to the dialog box, in place of the edit box. BIF_EDITBOX overrides this flag.
      BIF_UAHINT              :   100,
      // Version 6.0. Do not include the New Folder button in the browse dialog box.
      BIF_NONEWFOLDERBUTTON   :   200,
      // Version 6.0. When the selected item is a shortcut, return the PIDL of the shortcut itself rather than its target.
      BIF_NOTRANSLATETARGETS  :   400,
      // Only return computers. If the user selects anything other than a computer, the OK button is grayed.
      BIF_BROWSEFORCOMPUTER   :  1000,
      // Only allow the selection of printers. If the user selects anything other than a printer, the OK button is grayed. In Windows XP and later systems, the best practice is to use a Windows XP-style dialog, setting the root of the dialog to the Printers and Faxes folder (CSIDL_PRINTERS).
      BIF_BROWSEFORPRINTER    :  2000,
      // Version 4.71. The browse dialog box displays files as well as folders.
      BIF_BROWSEINCLUDEFILES  :  4000,
      // Version 5.0. The browse dialog box can display sharable resources on remote systems. This is intended for applications that want to expose remote shares on a local system. The BIF_NEWDIALOGSTYLE flag must also be set.
      BIF_SHAREABLE           :  8000,
      // Windows 7 and later. Allow folder junctions such as a library or a compressed file with a .zip file name extension to be browsed.
      BIF_BROWSEFILEJUNCTIONS : 10000
    }

    this.ssf = {
      // File system directory that corresponds to the user's non-localized Startup program group.
      ALTSTARTUP         : 29,
      // File system directory that serves as a common repository for application-specific data. A typical path is C:\Documents and Settings\username\Application Data.
      APPDATA            : 26,
      // Virtual folder that contains the objects in the user's Recycle Bin.
      BITBUCKET          : 10,
      // File system directory that corresponds to the non-localized Startup program group for all users. Valid only for Windows NT systems.
      COMMONALTSTARTUP   : 30,
      // Application data for all users. A typical path is C:\Documents and Settings\All Users\Application Data.
      COMMONAPPDATA      : 35,
      // File system directory that contains files and folders that appear on the desktop for all users. A typical path is C:\Documents and Settings\All Users\Desktop. Valid only for Windows NT systems.
      COMMONDESKTOPDIR   : 25,
      // File system directory that serves as a common repository for the favorite URLs shared by all users. Valid only for Windows NT systems.
      COMMONFAVORITES    : 31,
      // File system directory that contains the directories for the common program groups that appear on the Start menu for all users. A typical path is C:\Documents and Settings\All Users\Start Menu\Programs. Valid only for Windows NT systems.
      COMMONPROGRAMS     : 23,
      // File system directory that contains the programs and folders that appear on the Start menu for all users. A typical path is C:\Documents and Settings\All Users\Start Menu. Valid only for Windows NT systems.
      COMMONSTARTMENU    : 22,
      // File system directory that contains the programs that appear in the Startup folder for all users. A typical path is C:\Documents and Settings\All Users\Microsoft\Windows\Start Menu\Programs\StartUp. Valid only for Windows NT systems.
      COMMONSTARTUP      : 24,
      // Virtual folder that contains icons for the Control Panel applications.
      CONTROLS           :  3,
      // File system directory that serves as a common repository for Internet cookies. A typical path is C:\Documents and Settings\username\Application Data\Microsoft\Windows\Cookies.
      COOKIES            : 33,
      // Windows desktop—the virtual folder that is the root of the namespace.
      DESKTOP            :  0,
      // File system directory used to physically store the file objects that are displayed on the desktop. It is not to be confused with the desktop folder itself, which is a virtual folder. A typical path is C:\Documents and Settings\username\Desktop.
      DESKTOPDIRECTORY   : 16,
      // My Computer—the virtual folder that contains everything on the local computer: storage devices, printers, and Control Panel. This folder can also contain mapped network drives.
      DRIVES             : 17,
      // File system directory that serves as a common repository for the user's favorite URLs. A typical path is C:\Documents and Settings\username\Favorites.
      FAVORITES          :  6,
      // Virtual folder that contains installed fonts. A typical path is C:\Windows\Fonts.
      FONTS              : 20,
      // File system directory that serves as a common repository for Internet history items.
      HISTORY            : 34,
      // File system directory that serves as a common repository for temporary Internet files. A typical path is C:\Users\username\AppData\Local\Microsoft\Windows\Temporary Internet Files.
      INTERNETCACHE      : 32,
      // File system directory that serves as a data repository for local (non-roaming) applications. A typical path is C:\Users\username\AppData\Local.
      LOCALAPPDATA       : 28,
      // My Pictures folder. A typical path is C:\Users\username\Pictures.
      MYPICTURES         : 39,
      // A file system folder that contains any link objects in the My Network Places virtual folder. It is not the same as ssfNETWORK, which represents the network namespace root. A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Network Shortcuts.
      NETHOOD            : 19,
      // Network Neighborhood—the virtual folder that represents the root of the network namespace hierarchy.
      NETWORK            : 18,
      // File system directory that serves as a common repository for a user's documents. A typical path is C:\Users\username\Documents.
      PERSONAL           :  5,
      // Virtual folder that contains installed printers.
      PRINTERS           :  4,
      // File system directory that contains any link objects in the Printers virtual folder. A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Printer Shortcuts.
      PRINTHOOD          : 27,
      // User's profile folder.
      PROFILE            : 40,
      // Program Files folder. A typical path is C:\Program Files.
      PROGRAMFILES       : 38,
      // Program Files folder. A typical path is C:\Program Files, or C:\Program Files (X86) on a 64-bit computer.
      PROGRAMFILESx86    : 48,
      // File system directory that contains the user's program groups (which are also file system directories). A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs.
      PROGRAMS           :  2,
      // File system directory that contains the user's most recently used documents. A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Recent.
      RECENT             :  8,
      // File system directory that contains Send To menu items. A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\SendTo.
      SENDTO             :  9,
      // File system directory that contains Start menu items. A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Start Menu.
      STARTMENU          : 11,
      // File system directory that corresponds to the user's Startup program group. The system starts these programs whenever any user first logs into their profile after a reboot. A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\StartUp.
      STARTUP            :  7,
      // The System folder. A typical path is C:\Windows\System32.
      SYSTEM             : 37,
      // System folder. A typical path is C:\Windows\System32, or C:\Windows\Syswow32 on a 64-bit computer.
      SYSTEMx86          : 41,
      // File system directory that serves as a common repository for document templates.
      TEMPLATES          : 21,
      // Windows directory. This corresponds to the %windir% or %SystemRoot% environment variables. A typical path is C:\Windows.
      WINDOWS            : 36
    }
  }

  function Folder() {
    // http://msdn.microsoft.com/en-us/library/bb787868%28v=vs.85%29.aspx
    // <field name="ParentFolder" type="Folder">Contains the parent Folder object.</field>
    // <field name="Title" type="String">Contains the title of the folder.</field>
    this.ParentFolder = new Folder
    this.Title = ''
    this.CopyHere = function() {
      /// <summary>Copies an item or items to a folder.</summary>
      /// <param name=""></param>
    }
    this.GetDetailsOf = function() {
      /// <summary>Retrieves details about an item in a folder. For example, its size, type, or the time of its last modification.</summary>
      /// <param name=""></param>
    }
    this.Items = function() {
      /// <summary>Retrieves a FolderItems object that represents the collection of items in the folder.</summary>
      /// <returns type="FolderItems"/>
      return new FolderItems
    }
    this.MoveHere = function() {
      /// <summary>Moves an item or items to this folder.</summary>
      /// <param name=""></param>
    }
    this.NewFolder = function() {
      /// <summary>Creates a new folder.</summary>
      /// <param name=""></param>
    }
    this.ParseName = function() {
      /// <summary>Creates and returns a FolderItem object that represents a specified item.</summary>
      /// <param name=""></param>
    }
  }

  function FolderItem() {
    // http://msdn.microsoft.com/en-us/library/bb787810%28v=vs.85%29.aspx
    /// <field name="GetFolder" type="Folder">Contains the item's Folder object, if the item is a folder.</field>
    /// <field name="GetLink" type="Link">Contains the item's ShellLinkObject object, if the item is a shortcut.</field>
    /// <field name="IsBrowsable" type="Boolean">Indicates if the item can be hosted inside a browser or Windows Explorer frame.</field>
    /// <field name="IsFileSystem" type="Boolean">Indicates if the item is part of the file system.</field>
    /// <field name="IsFolder" type="Boolean">Indicates if the item is a folder.</field>
    /// <field name="IsLink" type="Boolean">Indicates whether the item is a shortcut.</field>
    /// <field name="ModifyDate" type="String">Sets or gets the date and time that a file was last modified. ModifyDate can be used to retrieve the date and time that a folder was last modified, but cannot set it. "01/01/1900 6:05:00 PM"</field>
    /// <field name="Name" type="String">Sets or gets the item's name.</field>
    /// <field name="Parent" type="Folder">Gets an object that represents the parent of the item.</field>
    /// <field name="Path" type="String">Contains the item's full path and name.</field>
    /// <field name="Size" type="Number">Contains the item's size.</field>
    /// <field name="Type" type="Number">Contains a string representation of the item's type.</field>
    this.GetFolder = new Folder
    this.GetLink = new Link
    this.IsBrowsable = false
    this.IsFileSystem = false
    this.IsFolder = false
    this.IsLink = false
    this.ModifyDate = ''
    this.Name = ''
    this.Parent = null
    this.Path = ''
    this.Size = 0
    this.Type = 0
    this.InvokeVerb = function(verb) {
      /// <summary>Executes a verb on the item.</summary>
      /// <param name="verb">
      /// A string that specifies the verb to be executed.
      /// It must be one of the values returned by the item's FolderItemVerb.Name property.
      /// If no verb is specified, the default verb will be invoked.
      /// Typically "open"
      /// </param>
    }
    this.Verbs = function() {
      /// <summary>Retrieves the item's FolderItemVerbs object. This object is the collection of verbs that can be executed on the item.</summary>
      /// <returns type="FolderItemVerbs"/>
      return new FolderItemVerbs
    }
  }

  function FolderItems() {
    // http://msdn.microsoft.com/en-us/library/bb787800%28v=vs.85%29.aspx
    this.Count = 0
    this.Item = function(index) {
      /// <summary>Retrieves the FolderItem object for a specified item in the collection.</summary>
      /// <param name="index">number, zero-based index of the item to retrieve. This value must be less than the value of the Count property.</param>
      /// <returns type="FolderItem"/>
      return new FolderItem
    }
  }

  function FolderItemVerb() {
    // http://msdn.microsoft.com/en-us/library/bb774172%28v=vs.85%29.aspx
    /// <field name="Name" type="String">Contains the verb's name.</field>
    this.Name = ''
    this.DoIt = function() {
      /// <summary>Executes a verb on the FolderItem associated with the verb.</summary>
    }
  }

  function FolderItemVerbs() {
    // http://msdn.microsoft.com/en-us/library/bb774160%28v=vs.85%29.aspx
    this.Count = 0
    this.Item = function(index) {
      /// <summary>Retrieves the FolderItem object for a specified item in the collection.</summary>
      /// <param name="index">number, zero-based index of the item to retrieve. This value must be less than the value of the Count property.</param>
      /// <returns type="FolderItemVerb"/>
      return new FolderItemVerb
    }
  }

  function Link() {
    // http://msdn.microsoft.com/en-us/library/bb774004%28v=vs.85%29
    // <field name="Arguments" type="String">Contains a link's arguments.</field>
    // <field name="Description" type="String">Gets or sets the description of the link.</field>
    // <field name="Hotkey" type="String">Gets or sets the keyboard shortcut for the link.</field>
    // <field name="Path" type="String">Gets or sets the path to the link object.</field>
    // <field name="ShowCommand" type="String">Gets or sets the initial display state (sized, minimized, or maximized) of the link's command.</field>
    // <field name="WorkingDirectory" type="String">Gets or sets the working directory specified in the link.</field>
    this.Arguments = ''
    this.Description = ''
    this.Hotkey = ''
    this.Path = ''
    this.ShowCommand = ''
    this.WorkingDirectory = ''
    this.GetIconLocation = function() {
      /// <summary>Gets the location of the icon assigned to the link.</summary>
    }
    this.Resolve = function() {
      /// <summary>Looks for the target of a Shell link, even if the target has been moved or renamed.</summary>
    }
    this.Save = function() {
      /// <summary>Saves all changes to the link.</summary>
    }
    this.SetIconLocation = function() {
      /// <summary>Sets the location of the icon assigned to the link.</summary>
    }
  }

}();

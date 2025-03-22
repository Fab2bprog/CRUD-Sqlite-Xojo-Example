#tag DesktopWindow
Begin DesktopWindow Win_Welcome
   Backdrop        =   0
   BackgroundColor =   &cFFFFFF
   Composite       =   False
   DefaultLocation =   2
   FullScreen      =   False
   HasBackgroundColor=   False
   HasCloseButton  =   True
   HasFullScreenButton=   False
   HasMaximizeButton=   True
   HasMinimizeButton=   True
   Height          =   400
   ImplicitInstance=   True
   MacProcID       =   0
   MaximumHeight   =   32000
   MaximumWidth    =   32000
   MenuBar         =   1840517119
   MenuBarVisible  =   False
   MinimumHeight   =   64
   MinimumWidth    =   64
   Resizeable      =   True
   Title           =   "CRUD Xojo Sqlite Database Example"
   Type            =   0
   Visible         =   True
   Width           =   674
   Begin DesktopImageViewer Image_LicenceMIT
      Active          =   False
      AllowAutoDeactivate=   True
      AllowTabStop    =   True
      Enabled         =   True
      Height          =   74
      Image           =   1607163903
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   20
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   False
      PanelIndex      =   0
      Scope           =   0
      TabIndex        =   0
      TabPanelIndex   =   0
      Tooltip         =   ""
      Top             =   299
      Transparent     =   False
      Visible         =   True
      Width           =   123
      _mIndex         =   0
      _mInitialParent =   ""
      _mName          =   ""
      _mPanelIndex    =   0
   End
   Begin DesktopLabel Lab_MITLicence
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   35
      Index           =   -2147483648
      Italic          =   False
      Left            =   147
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   False
      Multiline       =   True
      Scope           =   0
      Selectable      =   False
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "This software is licensed under the MIT (Massachusetts Institute of Technology) license. Author Fabrice Garcia, Borgo, Haute-Corse, France"
      TextAlignment   =   0
      TextColor       =   &c000000
      Tooltip         =   ""
      Top             =   291
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   500
   End
   Begin DesktopSeparator Separator1
      Active          =   False
      AllowAutoDeactivate=   True
      AllowTabStop    =   True
      Enabled         =   True
      Height          =   20
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   0
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   False
      PanelIndex      =   0
      Scope           =   0
      TabIndex        =   2
      TabPanelIndex   =   0
      Tooltip         =   ""
      Top             =   274
      Transparent     =   False
      Visible         =   True
      Width           =   674
      _mIndex         =   0
      _mInitialParent =   ""
      _mName          =   ""
      _mPanelIndex    =   0
   End
   Begin DesktopLabel Lab_Info
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   54
      Index           =   -2147483648
      Italic          =   False
      Left            =   147
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   False
      Multiline       =   True
      Scope           =   0
      Selectable      =   False
      TabIndex        =   3
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "The Goal is to provide you with an example showing you how to properly Create, Read, Modify, Delete information in a Sqlite database with Xojo. To do this, this program uses a data list and an input form."
      TextAlignment   =   0
      TextColor       =   &c000000
      Tooltip         =   ""
      Top             =   326
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   500
   End
   Begin DesktopLabel Labe_Info2
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   65
      Index           =   -2147483648
      Italic          =   False
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Multiline       =   True
      Scope           =   0
      Selectable      =   False
      TabIndex        =   4
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "Use the ""File"" menu to open an existing database ( database_example.sqlite ) provided with the example or to create a new database. Then select the ""Customer"" option from the ""Tables"" menu"
      TextAlignment   =   2
      TextColor       =   &c000000
      Tooltip         =   ""
      Top             =   100
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   634
   End
End
#tag EndDesktopWindow

#tag WindowCode
	#tag Event
		Sub Opening()
		  Dim MenuFile as Class_MenuWelcome
		  
		  // Me.Maximize
		  
		  DynaMenuBar = new DesktopMenuBar
		  Self.MenuBar   = DynaMenuBar
		  
		  Node_File = New DesktopMenuItem
		  Node_File.Text="File"
		  Node_File.Name="Node_File"
		  DynaMenuBar.AddMenu Node_File
		  
		  MenuFile = new Class_MenuWelcome
		  MenuFile.Node_Root=Node_File
		  MenuFile.Add_All
		End Sub
	#tag EndEvent


	#tag MenuHandler
		Function Ite_Article() As Boolean Handles Ite_Article.Action
		  // Win_ArticleLst.Show
		  Return True
		  
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function Ite_CloseBase() As Boolean Handles Ite_CloseBase.Action
		  Close_Base
		  
		  Return True
		  
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function Ite_Customer() As Boolean Handles Ite_Customer.Action
		  Win_CustomerList.ShowModal
		  Return True
		  
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function Ite_NewBase() As Boolean Handles Ite_NewBase.Action
		  // Creation d'une nouvelle base de donnée
		  
		  dim dlog as SaveFileDialog
		  dim file as FolderItem
		  
		  Var d As New MessageDialog                   ' declare the MessageDialog object
		  Var b As MessageDialogButton                 ' for handling the result
		  
		  
		  
		  if not (App.MainDB_State="CLOSE") then
		    
		    d.IconType = MessageDialog.IconTypes.Caution ' display warning icon
		    d.ActionButton.Caption = "Close"
		    d.CancelButton.Visible = True                ' show the Cancel button
		    d.AlternateActionButton.Visible = True      
		    d.AlternateActionButton.Caption = "Don't close"
		    d.Message = "A database is already open"
		    d.Explanation = "Do you want to close the current database? "
		    
		    b = d.ShowModal                             
		    Select Case b                                
		    Case d.ActionButton
		      Close_Base
		    Case d.AlternateActionButton
		      return false
		    Case d.CancelButton
		      return false
		    End Select
		    
		  end if
		  
		  
		  // Create a file creation dialog box
		  dlog = New SaveFileDialog
		  dlog.PromptText = "Creation of a new database"
		  dlog.SuggestedFileName = "mydatabase.sqlite"
		  file = dlog.ShowModal
		  
		  // In the case where the user cancels the choice of a file
		  if file = NIL then
		    return false
		  end
		  
		  // Delete a file with the same name
		  if file.Exists then
		    file.Remove
		    if file.Exists then
		      MessageBox "The File could not be replaced: this may be due either to a limitation of your user account, or to the fact that another Application is using this file"
		      return false
		    end if
		  end if
		  
		  App.MainDB = New SQLiteDatabase
		  App.MainDB.DatabaseFile= new FolderItem( file.NativePath, FolderItem.PathModes.Shell  )
		  
		  
		  if not App.CreateDatabaseFile then return false
		  
		  
		  
		  Self.Add_Menu
		  
		  
		  return true
		  
		  
		  
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function Ite_OpenBase() As Boolean Handles Ite_OpenBase.Action
		  // Ouverture d'une base de donnée
		  
		  dim dlog as OpenFileDialog
		  dim file as folderItem
		  
		  Var d As New MessageDialog                   ' declare the MessageDialog object
		  Var b As MessageDialogButton                 ' for handling the result
		  
		  Var SqliteType As New FileType
		  SqliteType.Name = "application/vnd.sqlite3"
		  SqliteType.Extensions = ".sqlite;.sqlite3;.db;.db3;.s3db;.sl3"
		  
		  if not (App.MainDB_State="CLOSE") then
		    
		    d.IconType = MessageDialog.IconTypes.Caution ' display warning icon
		    d.ActionButton.Caption = "Close"
		    d.CancelButton.Visible = True                ' show the Cancel button
		    d.AlternateActionButton.Visible = True      
		    d.AlternateActionButton.Caption = "Don't close"
		    d.Message = "A database is already open"
		    d.Explanation = "Do you want to close the current database? "
		    
		    b = d.ShowModal                             
		    Select Case b                                
		    Case d.ActionButton
		      Close_Base
		    Case d.AlternateActionButton
		      return false
		    Case d.CancelButton
		      return false
		    End Select
		    
		  end if
		  
		  // Create a file  dialog box
		  dlog = New OpenFileDialog
		  dlog.PromptText = "Opening an existing database"
		  dlog.SuggestedFileName = "mydatabase.sqlite"
		  dlog.Filter=SqliteType.Extensions
		  file = dlog.ShowModal
		  
		  // In the case where the user cancels the choice of a file
		  if file = NIL then
		    return false
		  end
		  
		  App.MainDB = New SQLiteDatabase
		  App.MainDB.databaseFile = New FolderItem( file.NativePath, FolderItem.PathModes.Shell)
		  
		  
		  if not App.OpenDatabaseFile then return false
		  
		  
		  Self.Init_Menu
		  Self.Add_Menu
		  
		  return true
		  
		  
		  
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function Ite_UnitMeasure() As Boolean Handles Ite_UnitMeasure.Action
		  // Win_UnitMesureLst.Show
		  Return True
		  
		End Function
	#tag EndMenuHandler


	#tag Method, Flags = &h0
		Sub Add_Menu()
		  Add_MenuTable
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Add_MenuTable()
		  Var MenuTable as Class_MenuTable
		  
		  Node_Table       = New DesktopMenuItem
		  Node_Table.text  = "Tables"
		  Node_Table.Name  = "Node_Table"
		  Node_Table.Enabled = true
		  DynaMenuBar.AddMenu Node_Table
		  
		  MenuTable = new Class_MenuTable
		  MenuTable.Node_Root=Node_Table
		  MenuTable.Add_Customers
		  
		  
		  
		  
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Close_Base()
		  App.CloseDatabaseFile
		  Init_Menu
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Init_Menu()
		  Var MenuFile as Class_MenuWelcome
		  
		  DynaMenuBar= new DesktopMenuBar
		  Self.MenuBar=DynaMenuBar
		  
		  Node_File = New DesktopMenuItem
		  Node_File.Text="File"
		  Node_File.Name="Node_File"
		  DynaMenuBar.AddMenu Node_File
		  
		  MenuFile  = new Class_MenuWelcome
		  MenuFile.Node_Root=Node_File
		  MenuFile.Add_All
		  
		  
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		DynaMenuBar As DesktopMenuBar
	#tag EndProperty

	#tag Property, Flags = &h0
		Node_File As DesktopMenuItem
	#tag EndProperty

	#tag Property, Flags = &h0
		Node_Table As DesktopMenuItem
	#tag EndProperty


#tag EndWindowCode

#tag ViewBehavior
	#tag ViewProperty
		Name="Name"
		Visible=true
		Group="ID"
		InitialValue=""
		Type="String"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Interfaces"
		Visible=true
		Group="ID"
		InitialValue=""
		Type="String"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Super"
		Visible=true
		Group="ID"
		InitialValue=""
		Type="String"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Width"
		Visible=true
		Group="Size"
		InitialValue="600"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Height"
		Visible=true
		Group="Size"
		InitialValue="400"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimumWidth"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimumHeight"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximumWidth"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximumHeight"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Type"
		Visible=true
		Group="Frame"
		InitialValue="0"
		Type="Types"
		EditorType="Enum"
		#tag EnumValues
			"0 - Document"
			"1 - Movable Modal"
			"2 - Modal Dialog"
			"3 - Floating Window"
			"4 - Plain Box"
			"5 - Shadowed Box"
			"6 - Rounded Window"
			"7 - Global Floating Window"
			"8 - Sheet Window"
			"9 - Modeless Dialog"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Title"
		Visible=true
		Group="Frame"
		InitialValue="Untitled"
		Type="String"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasCloseButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasMaximizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasMinimizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasFullScreenButton"
		Visible=true
		Group="Frame"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Resizeable"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Composite"
		Visible=false
		Group="OS X (Carbon)"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MacProcID"
		Visible=false
		Group="OS X (Carbon)"
		InitialValue="0"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="FullScreen"
		Visible=true
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="DefaultLocation"
		Visible=true
		Group="Behavior"
		InitialValue="2"
		Type="Locations"
		EditorType="Enum"
		#tag EnumValues
			"0 - Default"
			"1 - Parent Window"
			"2 - Main Screen"
			"3 - Parent Window Screen"
			"4 - Stagger"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Visible"
		Visible=true
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="ImplicitInstance"
		Visible=true
		Group="Window Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasBackgroundColor"
		Visible=true
		Group="Background"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="BackgroundColor"
		Visible=true
		Group="Background"
		InitialValue="&cFFFFFF"
		Type="ColorGroup"
		EditorType="ColorGroup"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Backdrop"
		Visible=true
		Group="Background"
		InitialValue=""
		Type="Picture"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBar"
		Visible=true
		Group="Menus"
		InitialValue=""
		Type="DesktopMenuBar"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBarVisible"
		Visible=true
		Group="Deprecated"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
#tag EndViewBehavior

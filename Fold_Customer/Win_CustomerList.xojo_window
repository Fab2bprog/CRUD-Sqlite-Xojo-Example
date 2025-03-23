#tag DesktopWindow
Begin DesktopWindow Win_CustomerList
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
   Height          =   446
   ImplicitInstance=   True
   MacProcID       =   0
   MaximumHeight   =   32000
   MaximumWidth    =   32000
   MenuBar         =   ""
   MenuBarVisible  =   False
   MinimumHeight   =   64
   MinimumWidth    =   64
   Resizeable      =   True
   Title           =   "Customers List"
   Type            =   0
   Visible         =   True
   Width           =   746
   Begin DesktopGroupBox Grb_Result
      AllowAutoDeactivate=   True
      Bold            =   False
      Caption         =   "Results"
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   343
      Index           =   -2147483648
      Italic          =   False
      Left            =   20
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   83
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   706
      Begin DesktopListBox LST_RESULT
         AllowAutoDeactivate=   True
         AllowAutoHideScrollbars=   True
         AllowExpandableRows=   False
         AllowFocusRing  =   True
         AllowResizableColumns=   False
         AllowRowDragging=   False
         AllowRowReordering=   False
         Bold            =   False
         ColumnCount     =   5
         ColumnWidths    =   ""
         DefaultRowHeight=   -1
         DropIndicatorVisible=   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         GridLineStyle   =   2
         HasBorder       =   True
         HasHeader       =   True
         HasHorizontalScrollbar=   False
         HasVerticalScrollbar=   True
         HeadingIndex    =   -1
         Height          =   319
         Index           =   -2147483648
         InitialParent   =   "Grb_Result"
         InitialValue    =   "Number	First_Name	Second_Name	Credit_Amount	BirthDay"
         Italic          =   False
         Left            =   31
         LockBottom      =   True
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   True
         LockTop         =   True
         RequiresSelection=   False
         RowSelectionType=   0
         Scope           =   0
         TabIndex        =   0
         TabPanelIndex   =   0
         TabStop         =   True
         Tooltip         =   ""
         Top             =   99
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   683
         _ScrollOffset   =   0
         _ScrollWidth    =   -1
      End
   End
   Begin Tolb_ToolList Tolb_CustomerLst
      Enabled         =   True
      Index           =   -2147483648
      LockedInPosition=   False
      Scope           =   0
      TabPanelIndex   =   0
      Visible         =   True
   End
   Begin DesktopGroupBox Grb_Search
      AllowAutoDeactivate=   True
      Bold            =   False
      Caption         =   "Search criteria"
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   68
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   11
      Transparent     =   True
      Underline       =   False
      Visible         =   True
      Width           =   706
      Begin DesktopLabel Lab_TypeField
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "Grb_Search"
         Italic          =   False
         Left            =   28
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   False
         LockRight       =   False
         LockTop         =   False
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   0
         TabPanelIndex   =   0
         TabStop         =   True
         Text            =   "Field"
         TextAlignment   =   0
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   36
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   47
      End
      Begin DesktopPopupMenu Pop_FieldName
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   25
         Index           =   -2147483648
         InitialParent   =   "Grb_Search"
         InitialValue    =   ""
         Italic          =   False
         Left            =   87
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         SelectedRowIndex=   0
         TabIndex        =   2
         TabPanelIndex   =   0
         TabStop         =   True
         Tooltip         =   ""
         Top             =   35
         Transparent     =   True
         Underline       =   False
         Visible         =   True
         Width           =   252
      End
      Begin DesktopPopupMenu Pop_Operator
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   25
         Index           =   -2147483648
         InitialParent   =   "Grb_Search"
         InitialValue    =   "EQUAL\nGREATER_EQUAL\nLESS_EQUAL\nDIFFERENT\nGREATER\nLOWER\nCONTAINS\nCONTAINS_NOT\nSTARTS_WITH\nFINISHED_BY\nIS_EMPTY\nIS_NOT_EMPTY\nIS_NULL\nIS_NOT_NULL"
         Italic          =   False
         Left            =   351
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   0
         SelectedRowIndex=   0
         TabIndex        =   3
         TabPanelIndex   =   0
         TabStop         =   True
         Tooltip         =   ""
         Top             =   35
         Transparent     =   True
         Underline       =   False
         Visible         =   True
         Width           =   145
      End
      Begin DesktopSearchField Search_FindData
         Active          =   False
         AllowAutoDeactivate=   True
         AllowFocusRing  =   False
         AllowRecentItems=   False
         AllowTabStop    =   True
         ClearMenuItemValue=   ""
         Enabled         =   True
         Height          =   22
         Hint            =   ""
         Index           =   -2147483648
         InitialParent   =   "Grb_Search"
         Left            =   508
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MaximumRecentItems=   5
         PanelIndex      =   0
         RecentItemsValue=   ""
         Scope           =   0
         TabIndex        =   4
         TabPanelIndex   =   0
         Text            =   ""
         Tooltip         =   ""
         Top             =   36
         Transparent     =   False
         Visible         =   True
         Width           =   115
         _mIndex         =   0
         _mInitialParent =   ""
         _mName          =   ""
         _mPanelIndex    =   0
      End
      Begin DesktopButton Button1
         AllowAutoDeactivate=   True
         Bold            =   False
         Cancel          =   False
         Caption         =   "OK"
         Default         =   True
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "Grb_Search"
         Italic          =   False
         Left            =   635
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MacButtonStyle  =   0
         Scope           =   0
         TabIndex        =   5
         TabPanelIndex   =   0
         TabStop         =   True
         Tooltip         =   ""
         Top             =   36
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   80
      End
   End
End
#tag EndDesktopWindow

#tag WindowCode
	#tag Event
		Sub Opening()
		  Var ColNameRS as RowSet
		  
		  SrcDataObject = new Class_Customer
		  SrcDataObject.DBaseID=App.MainDB
		  
		  // Constructs a query whose sole purpose is to collect the names of the fields in a table
		  ColNameRS = App.MainDB.SelectSQL( "SELECT name FROM PRAGMA_TABLE_INFO('"+SrcDataObject.TableName+"');")
		  
		  
		  // Allows you to automatically collect the list of name fields in a table and fill the drop-down menu with this information
		  Pop_FieldName.RemoveAllRows
		  While Not ColNameRS.AfterLastRow
		    Pop_FieldName.AddRow  ColNameRS.Column("name").StringValue
		    ColNameRS.MoveToNextRow
		  wend
		  Pop_FieldName.SelectedRowIndex=0
		  
		  
		  
		  
		  
		  // Runs a global search on the table
		  Action_SearchAll
		  
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h0
		Sub Action_Create()
		  // This method is used to call the customer form window in record create mode
		  
		  Win_customerForm.Show
		  Win_customerForm.Called_From="Win_CustomerList"
		  Call Win_customerForm.CrudModeSelect("CREATE")
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Action_Delete()
		  // This method allows the deletion of a line from the list and the record associated with it in the sqlite database
		  
		  Var MesDiag As New MessageDialog                   ' declare the MessageDialog object
		  Var MDiagB As MessageDialogButton                 ' for handling the result
		  
		  Var KeySelection As Int64
		  
		  
		  if LST_RESULT.SelectedRowCount =0 Then
		    MessageBox("To use this function you must first select an element in the list")
		    exit sub
		  end if
		  
		  MesDiag.IconType = MessageDialog.IconTypes.Caution ' display warning icon
		  MesDiag.ActionButton.Caption = "Yes"
		  MesDiag.CancelButton.Visible = True                ' show the Cancel button
		  MesDiag.AlternateActionButton.Visible = True      
		  MesDiag.AlternateActionButton.Caption = "No"
		  MesDiag.Message = "Removing an item from the database"
		  MesDiag.Explanation = "Are you sure you want to delete the selected customer? "
		  
		  MDiagB = MesDiag.ShowModal                             
		  Select Case MDiagB                                
		  Case MesDiag.ActionButton
		    
		    KeySelection= LST_RESULT.CellTextAt(LST_RESULT.SelectedRowIndex,0).ToInt64
		    if SrcDataObject.Record_Delete(KeySelection) then
		      //Remove the selected row
		      LST_RESULT.RemoveRowAt(LST_RESULT.SelectedRowIndex)
		    end
		    
		    
		  Case MesDiag.AlternateActionButton
		    exit sub
		  Case MesDiag.CancelButton
		    exit sub
		  End Select
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Action_Exit()
		  self.Close
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Action_ReadOnly()
		  // This method is used to call the customer form window in record read only mode
		  Var KeySelection as Int64
		  
		  
		  if LST_RESULT.SelectedRowCount=0 Then
		    MessageBox("To use this function you must first select an element in the list")
		    exit sub
		  end if
		  
		  
		  Win_customerForm.Show
		  Win_customerForm.Called_From="Win_CustomerList"
		  KeySelection= LST_RESULT.CellTextAt(LST_RESULT.SelectedRowIndex,0).ToInt64
		  
		  Call Win_customerForm.CrudModeSelect("READ", KeySelection)
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Action_Search()
		  // This method builds a parameterized SQL query with information defined by the user
		  
		  SrcDataObject.SqlQuerySource="SELECT *  FROM "+SrcDataObject.TableName+" WHERE "+  SearchCondition( )+" ORDER BY  "+SrcDataObject.KeyTableName
		  SrcDataObject.SqlParams.RemoveAll
		  
		  select case  Pop_Operator.SelectedRowText
		  Case "IS_EMPTY", "IS_NOT_EMPTY","IS_NULL","IS_NOT_NULL"
		    // Do nothing
		  Case "CONTAINS","CONTAINS_NOT"
		    SrcDataObject.SqlParams.Add("%"+Search_FindData.Text+"%")
		  Case "STARTS_WITH"
		    SrcDataObject.SqlParams.Add(Search_FindData.Text+"%")
		  Case "FINISHED_BY"
		    SrcDataObject.SqlParams.Add("%"+Search_FindData.Text)
		  else
		    SrcDataObject.SqlParams.Add(Search_FindData.Text)
		  end select
		  
		  List_FillData()
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Action_SearchAll()
		  // This method creates a query displaying all the records in the table without any selection criteria.
		  
		  SrcDataObject.SqlQuerySource="SELECT *  FROM "+SrcDataObject.TableName+" ORDER BY  "+SrcDataObject.KeyTableName
		  SrcDataObject.SqlParams.RemoveAll
		  
		  List_FillData()
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Action_Update()
		  // This method is used to call the customer form window in record modification mode
		  
		  Var KeySelection As Int64
		  
		  if LST_RESULT.SelectedRowCount =0 Then
		    MessageBox("To use this function you must first select an element in the list")
		    exit sub
		  end if
		  
		  Win_customerForm.Show
		  Win_customerForm.Called_From="Win_CustomerList"
		  KeySelection= LST_RESULT.CellTextAt(LST_RESULT.SelectedRowIndex,0).ToInt64
		  Call Win_customerForm.CrudModeSelect("UPDATE", KeySelection)
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub List_FillData(PreserveSelection as Boolean = false)
		  // This method allows you to fill the list from the data of a query
		  
		  Var SelRow as Int64 =-1
		  
		  // Memorize the selection
		  if LST_RESULT.SelectedRowCount>0 and PreserveSelection=true then
		    SelRow=LST_RESULT.SelectedRowIndex
		  end if
		  
		  // Delete All row
		  LST_RESULT.RemoveAllRows
		  
		  // Run the Query
		  if not SrcDataObject.Run_SqlQuerySource()    then exit sub
		  if not SrcDataObject .Record_ReadFirst()    then exit sub
		  
		  // Fill the list
		  Do
		    
		    LST_RESULT.AddRow  str(SrcDataObject.KeyTableValue)
		    LST_RESULT.CellTextAt(LST_RESULT.LastRowIndex,1) = str(SrcDataObject.Field_FirstName)
		    LST_RESULT.CellTextAt(LST_RESULT.LastRowIndex,2) = str(SrcDataObject.Field_LastName)
		    LST_RESULT.CellTextAt(LST_RESULT.LastRowIndex,3) = str(SrcDataObject.Field_CreditAmount)
		    LST_RESULT.CellTextAt(LST_RESULT.LastRowIndex,4) = str(SrcDataObject.Field_BirthDay.SQLDate)
		    
		    // Exits Do...Loop when we reach the end of the query records
		    if not SrcDataObject.Record_ReadNext then exit do
		    
		  Loop
		  
		  // ReSelects the row that had been selected previously when refreshing the table
		  if SelRow>=0 and PreserveSelection=true then
		    if LST_RESULT.RowCount<=SelRow then LST_RESULT.SelectedRowIndex=SelRow
		  end if
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function SearchCondition() As string
		  Var SqlString As String
		  Var OperatorSql as String
		  
		  
		  Var FieldValue as String ="?"
		  Var FieldName as String   = Pop_FieldName.SelectedRowText
		  
		  
		  Select Case Pop_Operator.SelectedRowText
		    
		  Case "EQUAL"
		    OperatorSql =" ="
		    
		  Case "GREATER_EQUAL"
		    OperatorSql =" >="
		    
		  Case "LESS_EQUAL"
		    OperatorSql =" <="
		    
		  Case "DIFFERENT"
		    OperatorSql =" <>"
		    
		  Case "GREATER"
		    OperatorSql =" >"
		    
		  Case "LOWER"
		    OperatorSql =" <"
		    
		  Case "CONTAINS"
		    OperatorSql =" LIKE "
		    
		  Case "CONTAINS_NOT"
		    OperatorSql =" NOT LIKE "
		    
		  Case "STARTS_WITH"
		    OperatorSql =" LIKE "
		    
		  Case "FINISHED_BY"
		    OperatorSql =" LIKE "
		    
		  Case "IS_EMPTY"
		    OperatorSql =" = "
		    FieldValue="''"
		    
		  Case "IS_NOT_EMPTY"
		    OperatorSql =" <> "
		    FieldValue="''"
		    
		  Case "IS_NULL"
		    OperatorSql =" IS NULL "
		    FieldValue=""
		    
		  Case "IS_NOT_NULL"
		    OperatorSql =" IS NOT NULL "
		    FieldValue=""
		    
		  End select
		  
		  SqlString = FieldName + OperatorSql +FieldValue
		  
		  return SqlString
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function Validation_Key(key as String) As Boolean
		  if key=chr(13)  then return true
		  if key=chr(3)    then return true
		End Function
	#tag EndMethod


	#tag Property, Flags = &h21
		#tag Note
			// Represents the data source class which will allow you to read, create and modify your data in the database
		#tag EndNote
		Private SrcDataObject As Class_Customer
	#tag EndProperty


#tag EndWindowCode

#tag Events LST_RESULT
	#tag Event
		Function PaintCellBackground(g As Graphics, row As Integer, column As Integer) As Boolean
		  if row mod 2 = 0 then
		    g.DrawingColor = &cFFFFFF // white color
		  else
		    g.DrawingColor = &cf3f7fb // blue color
		  end if
		  g.FillRectangle 0,0, g.Width, g.Height
		End Function
	#tag EndEvent
	#tag Event
		Sub DoublePressed()
		  Action_Update
		End Sub
	#tag EndEvent
	#tag Event
		Function KeyDown(key As String) As Boolean
		  
		  if  Validation_Key(Key) then
		    Action_Update
		  end if
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events Tolb_CustomerLst
	#tag Event
		Sub Pressed(item As DesktopToolbarItem)
		  Select case item.Name
		    
		  Case "ToIt_Search"
		    // Launch a parameterized search
		    Call Action_Search
		  Case "ToIt_AllRecord"
		    // Launch a global search, show all record
		    Call Action_SearchAll
		  Case "ToIt_ReadOnly"
		    // Open the customer form in reading mode
		    Call Action_ReadOnly
		  Case "ToIt_Create"
		    // Open a blank customer form in creation mode for entering a new record
		    Call Action_Create
		  Case "ToIt_Update"
		    // Opens a customer form for data modification
		    Call Action_Update
		  Case "ToIt_Delete"
		    // Delete the selected line in the list
		    Call Action_Delete
		    
		  end select
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events Pop_FieldName
	#tag Event
		Function KeyDown(key As String) As Boolean
		  if Validation_Key(Key) then
		    FocusNext
		    return True
		  end if
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events Pop_Operator
	#tag Event
		Function KeyDown(key As String) As Boolean
		  if Validation_Key(Key) then
		    FocusNext
		    return True
		  end if
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events Search_FindData
	#tag Event
		Function KeyDown(key As String) As Boolean
		  if Validation_Key(Key) then
		    Action_Search
		    return true
		  end if
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events Button1
	#tag Event
		Sub Pressed()
		  Var i as int64
		  
		  For i = 1 to 5000
		    SrcDataObject.Field_FirstName=System.Random.InRange(1000,10000).ToString.ReplaceAll("1","A").ReplaceAll("2","B").ReplaceAll("3","C")
		    SrcDataObject.Field_LastName=System.Random.InRange(1000,10000).ToString.ReplaceAll("1","A").ReplaceAll("2","B").ReplaceAll("3","C")
		    SrcDataObject.Field_CreditAmount= System.Random.InRange(1000,10000)
		    SrcDataObject.Field_BirthDay=DateTime.Now
		    Call SrcDataObject.Record_Create
		  next
		  
		End Sub
	#tag EndEvent
#tag EndEvents
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

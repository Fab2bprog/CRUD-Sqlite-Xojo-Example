#tag DesktopWindow
Begin DesktopWindow Win_customerForm
   Backdrop        =   0
   BackgroundColor =   &cFFFFFF
   Composite       =   False
   DefaultLocation =   2
   FullScreen      =   False
   HasBackgroundColor=   False
   HasCloseButton  =   True
   HasFullScreenButton=   False
   HasMaximizeButton=   False
   HasMinimizeButton=   True
   Height          =   325
   ImplicitInstance=   True
   MacProcID       =   0
   MaximumHeight   =   32000
   MaximumWidth    =   32000
   MenuBar         =   609568767
   MenuBarVisible  =   False
   MinimumHeight   =   64
   MinimumWidth    =   64
   Resizeable      =   True
   Title           =   "Customer Form"
   Type            =   0
   Visible         =   True
   Width           =   520
   Begin DesktopGroupBox Grb_Identification
      AllowAutoDeactivate=   True
      Bold            =   False
      Caption         =   "Identification"
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   245
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   29
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Scope           =   0
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   "Valider"
      Top             =   20
      Transparent     =   True
      Underline       =   False
      Visible         =   True
      Width           =   471
      Begin DesktopTextField Txt_Code
         AllowAutoDeactivate=   True
         AllowFocusRing  =   True
         AllowSpellChecking=   False
         AllowTabs       =   False
         BackgroundColor =   &cFFFFFF00
         Bold            =   False
         Enabled         =   False
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Format          =   ""
         HasBorder       =   True
         Height          =   22
         Hint            =   ""
         Index           =   -2147483648
         InitialParent   =   "Grb_Identification"
         Italic          =   False
         Left            =   161
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         MaximumCharactersAllowed=   15
         Password        =   False
         ReadOnly        =   False
         Scope           =   0
         TabIndex        =   0
         TabPanelIndex   =   0
         TabStop         =   True
         Text            =   ""
         TextAlignment   =   0
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   47
         Transparent     =   True
         Underline       =   False
         ValidationMask  =   ""
         Visible         =   True
         Width           =   109
      End
      Begin DesktopLabel Lab_Code
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "Grb_Identification"
         Italic          =   False
         Left            =   41
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   False
         LockRight       =   False
         LockTop         =   False
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   1
         TabPanelIndex   =   0
         TabStop         =   True
         Text            =   "ID"
         TextAlignment   =   0
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   49
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   71
      End
      Begin DesktopLabel Lab_FirstName
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "Grb_Identification"
         Italic          =   False
         Left            =   41
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   False
         LockRight       =   False
         LockTop         =   False
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   2
         TabPanelIndex   =   0
         TabStop         =   True
         Text            =   "First Name"
         TextAlignment   =   0
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   89
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   108
      End
      Begin DesktopTextField Txt_FirstName
         AllowAutoDeactivate=   True
         AllowFocusRing  =   True
         AllowSpellChecking=   False
         AllowTabs       =   False
         BackgroundColor =   &cFFFFFF00
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Format          =   ""
         HasBorder       =   True
         Height          =   22
         Hint            =   ""
         Index           =   -2147483648
         InitialParent   =   "Grb_Identification"
         Italic          =   False
         Left            =   161
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   True
         LockTop         =   True
         MaximumCharactersAllowed=   50
         Password        =   False
         ReadOnly        =   False
         Scope           =   0
         TabIndex        =   3
         TabPanelIndex   =   0
         TabStop         =   True
         Text            =   ""
         TextAlignment   =   0
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   87
         Transparent     =   True
         Underline       =   False
         ValidationMask  =   ""
         Visible         =   True
         Width           =   320
      End
      Begin DesktopLabel Lab_LastName
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "Grb_Identification"
         Italic          =   False
         Left            =   41
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   False
         LockRight       =   False
         LockTop         =   False
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   4
         TabPanelIndex   =   0
         TabStop         =   True
         Text            =   "Last Name"
         TextAlignment   =   0
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   123
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   108
      End
      Begin DesktopTextField Txt_LastName
         AllowAutoDeactivate=   True
         AllowFocusRing  =   True
         AllowSpellChecking=   False
         AllowTabs       =   False
         BackgroundColor =   &cFFFFFF00
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Format          =   ""
         HasBorder       =   True
         Height          =   22
         Hint            =   ""
         Index           =   -2147483648
         InitialParent   =   "Grb_Identification"
         Italic          =   False
         Left            =   161
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   True
         LockTop         =   True
         MaximumCharactersAllowed=   50
         Password        =   False
         ReadOnly        =   False
         Scope           =   0
         TabIndex        =   5
         TabPanelIndex   =   0
         TabStop         =   True
         Text            =   ""
         TextAlignment   =   0
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   121
         Transparent     =   True
         Underline       =   False
         ValidationMask  =   ""
         Visible         =   True
         Width           =   320
      End
      Begin DesktopLabel Lab_Birthday
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "Grb_Identification"
         Italic          =   False
         Left            =   41
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   False
         LockRight       =   False
         LockTop         =   False
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   6
         TabPanelIndex   =   0
         TabStop         =   True
         Text            =   "BirthDay"
         TextAlignment   =   0
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   157
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   108
      End
      Begin DesktopLabel Lab_CreditAmount
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "Grb_Identification"
         Italic          =   False
         Left            =   41
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   False
         LockRight       =   False
         LockTop         =   False
         Multiline       =   False
         Scope           =   0
         Selectable      =   False
         TabIndex        =   8
         TabPanelIndex   =   0
         TabStop         =   True
         Text            =   "Credit Amount"
         TextAlignment   =   0
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   191
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   108
      End
      Begin DesktopTextField Txt_CreditAmount
         AllowAutoDeactivate=   True
         AllowFocusRing  =   True
         AllowSpellChecking=   False
         AllowTabs       =   False
         BackgroundColor =   &cFFFFFF00
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Format          =   ""
         HasBorder       =   True
         Height          =   22
         Hint            =   ""
         Index           =   -2147483648
         InitialParent   =   "Grb_Identification"
         Italic          =   False
         Left            =   161
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   True
         LockTop         =   True
         MaximumCharactersAllowed=   50
         Password        =   False
         ReadOnly        =   False
         Scope           =   0
         TabIndex        =   9
         TabPanelIndex   =   0
         TabStop         =   True
         Text            =   ""
         TextAlignment   =   0
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   191
         Transparent     =   True
         Underline       =   False
         ValidationMask  =   ""
         Visible         =   True
         Width           =   109
      End
      Begin DesktopDateTimePicker Dtp_BirthDay
         Active          =   False
         AllowAutoDeactivate=   True
         AllowFocusRing  =   False
         AllowTabStop    =   True
         DisplayMode     =   1
         DisplaySeconds  =   False
         Enabled         =   True
         GraphicalDisplay=   False
         Height          =   22
         HourMode        =   2
         Index           =   -2147483648
         InitialParent   =   "Grb_Identification"
         Left            =   161
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         PanelIndex      =   0
         Scope           =   0
         TabIndex        =   10
         TabPanelIndex   =   0
         TodayButtonCaption=   ""
         Tooltip         =   ""
         Top             =   157
         Transparent     =   False
         Visible         =   True
         Width           =   109
         _mIndex         =   0
         _mInitialParent =   ""
         _mName          =   ""
         _mPanelIndex    =   0
      End
   End
   Begin DesktopButton BTN_OK
      AllowAutoDeactivate=   True
      Bold            =   False
      Cancel          =   False
      Caption         =   "OK"
      Default         =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   22
      Index           =   -2147483648
      Italic          =   False
      Left            =   328
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MacButtonStyle  =   0
      Scope           =   0
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   283
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   80
   End
   Begin DesktopButton BTN_Cancel
      AllowAutoDeactivate=   True
      Bold            =   False
      Cancel          =   False
      Caption         =   "Cancel"
      Default         =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   22
      Index           =   -2147483648
      Italic          =   False
      Left            =   420
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MacButtonStyle  =   0
      Scope           =   0
      TabIndex        =   2
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   283
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   80
   End
End
#tag EndDesktopWindow

#tag WindowCode
	#tag Event
		Function CancelClosing(appQuitting As Boolean) As Boolean
		  if ForceExit then
		    // Returning false confirms that the window has been closed
		    return false
		  end if
		  
		  
		  if Confirm_UserExit then
		    // Returning false confirms that the window has been closed
		    return false
		  end if
		  
		  // Returning true cancels closing the window
		  return true
		  
		End Function
	#tag EndEvent


	#tag MenuHandler
		Function Ite_Exit() As Boolean Handles Ite_Exit.Action
		  self.close
		  return true
		  
		  
		End Function
	#tag EndMenuHandler


	#tag Method, Flags = &h0
		Function Confirm_UserExit() As Boolean
		  // Allow the user to confirm that they want to exit the form, in particular without saving the data they have entered
		  
		  Var d As New MessageDialog                   ' declare the MessageDialog object
		  Var b As MessageDialogButton                 ' for handling the result
		  
		  
		  
		  if  not Detect_Modif then return true
		  
		  d.IconType = MessageDialog.IconTypes.Caution ' display warning icon
		  d.ActionButton.Caption = "Yes"
		  d.CancelButton.Visible = True                ' show the Cancel button
		  d.AlternateActionButton.Visible = True      
		  d.AlternateActionButton.Caption = "No"
		  d.Message = "Data Entry detected"
		  d.Explanation = "Abandon the ongoing data entry ? "
		  
		  b = d.ShowModal                             
		  Select Case b                                
		  Case d.ActionButton
		    return true
		  Case d.AlternateActionButton
		    return false
		  Case d.CancelButton
		    return false
		  End Select
		  
		  return false
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ControlsToScreen()
		  // Updates the form fields relative to the data source
		  
		  Txt_Code.Text                       =    SrcDataObject.KeyTableValue.ToString
		  Txt_FirstName.Text               =    SrcDataObject.Field_FirstName
		  Txt_LastName.Text               =    SrcDataObject.Field_LastName
		  Txt_CreditAmount.Text         =    SrcDataObject.Field_CreditAmount.ToString
		  Dtp_BirthDay.SelectedDate =    SrcDataObject.Field_BirthDay
		  
		  Define_InitialValues
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CrudModeSelect(DataAccessMode as String, KeyDataValue as Int64 = -1) As Boolean
		  // This method must be called when the window is opened before its use by the user of the program
		  // The goal is to define how the form should behave, whether it involves modifying, creating or reading data.
		  
		  CrudMode=DataAccessMode
		  
		  if CrudMode="CREATE" then
		    
		    Lab_Code.Visible=False
		    Txt_Code.Visible=False
		    
		    SrcDataObject= new Class_Customer
		    SrcDataObject.DBaseID=App.MainDB
		    
		    return true
		    
		  end if
		  
		  if CrudMode="READ" then
		    
		    Enable_Form(False)
		    BTN_OK.Visible=false
		    
		  end if
		  
		  if CrudMode="UPDATE" then
		    
		    Txt_Code.ReadOnly=True
		    
		  end if
		  
		  
		  
		  SrcDataObject= new Class_Customer
		  SrcDataObject.DBaseID=App.MainDB
		  SrcDataObject.SqlQuerySource="SELECT * FROM "+SrcDataObject.TableName+" WHERE "+SrcDataObject.KeyTableName+"="+KeyDataValue.ToString
		  
		  if not SrcDataObject.Run_SqlQuerySource then return false
		  Call SrcDataObject.Record_ReadFirst
		  
		  
		  ControlsToScreen
		  
		  return true
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Define_InitialValues()
		  //Memorization of variables in their initial state in order to detect modifications to the form later
		  
		  Initial_Code                =  SrcDataObject.KeyTableValue.ToString
		  Initial_FirstName        =  SrcDataObject.Field_FirstName
		  Initial_LastName        =  SrcDataObject.Field_LastName
		  Initial_CreditAmount  =  SrcDataObject.Field_CreditAmount
		  Initial_Birthday           =  SrcDataObject.Field_BirthDay
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1000
		Function Detect_Modif() As Boolean
		  // Allows you to detect if the user has modified information in the form
		  
		  if not ( Txt_Code.Text                =  Initial_Code          )  then return true
		  if not ( Txt_FirstName.Text         =  Initial_FirstName    )  then return true
		  if not ( Txt_LastName.Text         =  Initial_LastName    )  then return true
		  if not ( Txt_CreditAmount.Text   =  Initial_CreditAmount.ToString    )  then return true
		  if not ( Dtp_BirthDay.SelectedDate  =  Initial_Birthday  )  then return true
		  
		  return false
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Enable_Form(TrueFalse as Boolean)
		  Grb_Identification.Enabled  = TrueFalse
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ScreenToControls()
		  // Transfer information from form fields to data source properties
		  
		  SrcDataObject.KeyTableValue        = Txt_Code.Text.ToInt64
		  SrcDataObject.Field_FirstName      = Txt_FirstName.Text
		  SrcDataObject.Field_LastName      = Txt_LastName.Text
		  SrcDataObject.Field_CreditAmount = Txt_CreditAmount.Text.ToDouble
		  SrcDataObject.Field_BirthDay         =Dtp_BirthDay.SelectedDate
		  
		  Define_InitialValues
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Touch_Valid(key as String) As Boolean
		  if key=chr(13)  then return true
		  if key=chr(3)    then return true
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Valid_Form() As Boolean
		  // Allows you to define the validation conditions of the form, ensures that the user has completed the fields correctly
		  
		  if Txt_FirstName.Text.Trim="" then
		    MessageBox ("The First Name field must be completed")
		    return false
		  end if
		  
		  return true
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		#tag Note
			//  Used to identify from which window the opening of this form was triggered
		#tag EndNote
		Called_From As String
	#tag EndProperty

	#tag Property, Flags = &h1
		#tag Note
			// ONLY 3 values possible : CREATE, READ, UPDATE
		#tag EndNote
		Protected CrudMode As String = "MOD"
	#tag EndProperty

	#tag Property, Flags = &h21
		Private ForceExit As Boolean = false
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// Store the initial value of the table field BirthDay
			// Allows you to detect changes in the form
		#tag EndNote
		Initial_Birthday As DateTime
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// Store the initial value of the table key
			// Allows you to detect changes in the form
		#tag EndNote
		Initial_Code As String
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			//  Store the initial value of the table field CreditAmount
			// Allows you to detect changes in the form
		#tag EndNote
		Initial_CreditAmount As Double
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// Store the initial value of the table field FirstName
			// Allows you to detect changes in the form
		#tag EndNote
		Initial_FirstName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// Store the initial value of the table field LastName
			// Allows you to detect changes in the form
		#tag EndNote
		Initial_LastName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// Represents the data source class which will allow you to read, create and modify your data in the database
		#tag EndNote
		SrcDataObject As Class_Customer
	#tag EndProperty


#tag EndWindowCode

#tag Events Txt_Code
	#tag Event
		Function KeyDown(key As String) As Boolean
		  if Touch_Valid(Key) then
		    FocusNext
		    return True
		  end if
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events Txt_FirstName
	#tag Event
		Function KeyDown(key As String) As Boolean
		  if Touch_Valid(Key) then
		    FocusNext
		    return True
		  end if
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events Txt_LastName
	#tag Event
		Function KeyDown(key As String) As Boolean
		  if Touch_Valid(Key) then
		    FocusNext
		    return True
		  end if
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events Txt_CreditAmount
	#tag Event
		Function KeyDown(key As String) As Boolean
		  // The goal is to only allow the entry of monetary-type decimal numbers
		  
		  return App.Valide_TxtNumKeydown(me,Key,false)
		End Function
	#tag EndEvent
	#tag Event
		Sub FocusLost()
		  // The goal is to only allow the entry of monetary-type decimal numbers
		  
		  Call App.Valide_TxtNumLostFocus(me,false,2,0,999999999,0) 
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events BTN_OK
	#tag Event
		Sub Pressed()
		  
		  if Valid_Form then
		    
		    ScreenToControls
		    
		    If CrudMode="UPDATE"     then call SrcDataObject.Record_Update
		    If CrudMode="CREATE"   then call SrcDataObject.Record_Create
		    
		    
		    if Called_From="Win_CustomerList" then
		      Call Win_CustomerList.List_FillData(True)
		    end if
		    
		    ForceExit=true
		    self.close
		    
		  end if
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events BTN_Cancel
	#tag Event
		Sub Pressed()
		  self.close
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
	#tag ViewProperty
		Name="Initial_FirstName"
		Visible=false
		Group="Behavior"
		InitialValue=""
		Type="String"
		EditorType="MultiLineEditor"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Initial_Code"
		Visible=false
		Group="Behavior"
		InitialValue=""
		Type="String"
		EditorType="MultiLineEditor"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Initial_CreditAmount"
		Visible=false
		Group="Behavior"
		InitialValue=""
		Type="Double"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Initial_LastName"
		Visible=false
		Group="Behavior"
		InitialValue=""
		Type="String"
		EditorType="MultiLineEditor"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Called_From"
		Visible=false
		Group="Behavior"
		InitialValue=""
		Type="String"
		EditorType="MultiLineEditor"
	#tag EndViewProperty
#tag EndViewBehavior

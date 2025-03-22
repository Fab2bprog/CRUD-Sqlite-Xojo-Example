#tag Class
Protected Class Class_MenuWelcome
	#tag Method, Flags = &h0
		Sub Add_All()
		  Add_New
		  Add_Open
		  Add_Close
		  Add_Separator
		  Add_Exit
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Add_Close()
		  Ite_CloseBase          = New DesktopMenuItem
		  Ite_CloseBase.text     = "Close file"
		  Ite_CloseBase.Name     = "Ite_CloseBase"
		  Ite_CloseBase.Enabled=True
		  Node_Root.AddMenu Ite_CloseBase
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Add_Exit()
		  Ite_FileQuit         = New DesktopQuitMenuItem
		  Ite_FileQuit.Text= "Exit"
		  Node_Root.AddMenu Ite_FileQuit
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Add_New()
		  Ite_NewBase       = New DesktopMenuItem
		  Ite_NewBase.text  = "New File"
		  Ite_NewBase.Name  = "Ite_NewBase"
		  Ite_NewBase.Enabled=True
		  Node_Root.AddMenu Ite_NewBase
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Add_Open()
		  Ite_OpenBase            = New DesktopMenuItem
		  Ite_OpenBase.text       = "Open File"
		  Ite_OpenBase.Name       = "Ite_OpenBase"
		  Ite_OpenBase.Enabled=true
		  Node_Root.AddMenu Ite_OpenBase
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Add_Separator()
		  Dim Ite_MenuSepare  as DesktopMenuItem
		  
		  Ite_MenuSepare              = New DesktopMenuItem
		  Ite_MenuSepare.text       = "-"
		  Ite_MenuSepare.Name       = "Ite_MenuSepare"
		  Node_Root.AddMenu Ite_MenuSepare
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		Ite_CloseBase As DesktopMenuItem
	#tag EndProperty

	#tag Property, Flags = &h0
		Ite_FileQuit As DesktopMenuItem
	#tag EndProperty

	#tag Property, Flags = &h0
		Ite_NewBase As DesktopMenuItem
	#tag EndProperty

	#tag Property, Flags = &h0
		Ite_OpenBase As DesktopMenuItem
	#tag EndProperty

	#tag Property, Flags = &h0
		Node_Root As DesktopMenuItem
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
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
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass

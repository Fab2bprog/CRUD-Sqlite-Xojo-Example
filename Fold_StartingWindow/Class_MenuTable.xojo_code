#tag Class
Protected Class Class_MenuTable
	#tag Method, Flags = &h0
		Sub Add_All()
		  Add_Customers
		  
		  
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Add_Customers()
		  Ite_Customer         = New DesktopMenuItem
		  Ite_Customer.text    = "Customer"
		  Ite_Customer.Name    = "Ite_Customer"
		  Ite_Customer.Enabled = true
		  Node_Root.AddMenu Ite_Customer
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		Ite_Customer As DesktopMenuItem
	#tag EndProperty

	#tag Property, Flags = &h0
		Node_Action As DesktopMenuItem
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

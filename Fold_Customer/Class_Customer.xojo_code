#tag Class
Protected Class Class_Customer
	#tag Method, Flags = &h0
		Sub Constructor()
		  Fields_Init
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub Fields_Init()
		  // Initializes class properties with default values
		  
		  // [Update-XojoCode-NextLines]
		  KeyTableValue         = 0
		  Field_FirstName       = ""
		  Field_LastName       =  ""
		  Field_CreditAmount  = 0
		  Field_BirthDay          = DateTime.Now
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub Fields_Load()
		  // Loads the values ​​of a table record into the corresponding properties of the class.
		  
		  // [Update-XojoCode-NextLines]
		  Try
		    KeyTableValue         = DBaseRS.Column(KeyTableName).Int64Value
		  Catch error As NilObjectException
		    KeyTableValue    = 0
		  End Try
		  
		  Try
		    Field_FirstName       = DBaseRS.Column("FIRSTNAME").StringValue
		  Catch error As NilObjectException
		    Field_FirstName    = ""
		  End Try
		  
		  Try
		    Field_LastName        = DBaseRS.Column("LASTNAME").StringValue
		  Catch error As NilObjectException
		    Field_LastName    = ""
		  End Try
		  
		  Try
		    Field_CreditAmount  = DBaseRS.Column("CREDITAMOUNT").DoubleValue
		  Catch error As NilObjectException
		    Field_CreditAmount    = 0
		  End Try
		  
		  Try
		    // Safe retrieval of the BIRTHDAY column value
		    Var birthDayCol As DatabaseColumn = DBaseRS.Column("BIRTHDAY")
		    
		    If birthDayCol <> Nil And birthDayCol.DateTimeValue <> Nil Then
		      // Column exists and has valid datetime value
		      Field_BirthDay = birthDayCol.DateTimeValue
		    Else
		      Var Date_Default As New DateTime(1970, 1, 1)
		      Field_BirthDay=Date_Default 
		    End If
		    
		  Catch error As NilObjectException
		    Var Date_Default As New DateTime(1970, 1, 1)
		    Field_BirthDay=Date_Default 
		  End Try
		  
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Find_FreeKeyTableValue() As int64
		  // The purpose of this function is to generate a primary key value for the table record
		  
		  Dim rs         as RowSet
		  
		  rs = DBaseID.SelectSQL("SELECT IFNULL(MAX("+KeyTableName+")+1,1) AS MAX_KEY_VALUE FROM "+TableName)
		  
		  if  not (rs=NIL) then
		    rs.MoveToFirstRow
		    KeyTableValue        = rs.Column("MAX_KEY_VALUE").Value
		  else
		    KeyTableValue        = 1
		  end if
		  
		  return KeyTableValue
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Record_Create(Create_KeyTableValue as Boolean = True) As Boolean
		  // Create a new record with class properties values.
		  
		  // By default the key is incremented, but you can force the key to a given value by update KeyTableValue propertie.
		  if Create_KeyTableValue= True then KeyTableValue=Find_FreeKeyTableValue()
		  
		  
		  Var row As New DatabaseRow
		  
		  // [Update-XojoCode-NextLines]
		  row.Column(KeyTableName).Int64Value  = KeyTableValue
		  row.Column("FIRSTNAME").StringValue  = Field_FirstName
		  row.Column("LASTNAME").StringValue   = Field_LastName
		  row.Column("CREDITAMOUNT").DoubleValue   = Field_CreditAmount
		  row.Column("BIRTHDAY").DateTimeValue  = Field_BirthDay
		  
		  Try
		    DBaseID.AddRow(TableName, row)
		  Catch error As DatabaseException
		    MessageBox("DB Error: " + error.Message)
		    return false
		  End Try
		  
		  return true
		  
		  
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Record_Delete(Del_KeySelection as int64 = -1) As Boolean
		  
		  Select Case Del_KeySelection
		    
		  Case Is >=0
		    
		    // if Del_KeySelection>=0 then you want to delete a specific record and not the current record
		    Try
		      DBaseID.ExecuteSQL("DELETE FROM "+TableName+" WHERE "+KeyTableName+" = "+Del_KeySelection.ToString)
		      return true
		    Catch error As DatabaseException
		      MessageBox("DB Error: " + error.Message)
		      return false
		    End Try
		    
		  Case else
		    
		    // if Del_KeySelection=-1 then Delete current record 
		    Try
		      DBaseRS.RemoveRow
		      return true
		    Catch error As DatabaseException
		      MessageBox("DB Error: " + error.Message)
		      return false
		    End Try
		    
		  end Select
		  
		  
		  return false
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Record_ReadFirst() As Boolean
		  // Read the first record of the sql query
		  
		  Fields_Init
		  
		  
		  Try
		    DBaseRS.MoveToFirstRow
		  Catch error As DatabaseException
		    MessageBox("DB Error: " + error.Message)
		    return false
		  End Try
		  
		  if  DBaseRS.BeforeFirstRow or DBaseRS.AfterLastRow  then return false
		  
		  
		  Fields_Load()
		  
		  return true
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Record_ReadNext() As Boolean
		  // Read the next record of the sql query
		  
		  Fields_Init
		  
		  Try
		    DBaseRS.MoveToNextRow
		  Catch error As DatabaseException
		    MessageBox("DB Error: " + error.Message)
		    return false
		  End Try
		  
		  if  DBaseRS.AfterLastRow  then return false
		  
		  
		  Fields_Load()
		  
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Record_ReadPrevious() As Boolean
		  // Read the previous record of the sql query
		  
		  Fields_Init
		  
		  Try
		    DBaseRS.MoveToPreviousRow
		  Catch error As DatabaseException
		    MessageBox("DB Error: " + error.Message)
		    return false
		  End Try
		  
		  if  DBaseRS.BeforeFirstRow  then return false
		  
		  
		  Fields_Load()
		  
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Record_Update() As Boolean
		  // Update a record with class properties
		  
		  Try
		    DBaseRS.EditRow
		    
		    // [Update-XojoCode-NextLines]
		    DBaseRS.Column(KeyTableName).Int64Value  = KeyTableValue
		    DBaseRS.Column("FIRSTNAME").StringValue  = Field_FirstName
		    DBaseRS.Column("LASTNAME").StringValue   = Field_LastName
		    DBaseRS.Column("CREDITAMOUNT").DoubleValue   = Field_CreditAmount
		    DBaseRS.Column("BIRTHDAY").DateTimeValue  = Field_BirthDay
		    
		    DBaseRS.SaveRow
		    
		  Catch error As DatabaseException
		    MessageBox("DB Error: " + error.Message)
		    return false
		  End Try
		  
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Run_SqlQuerySource() As Boolean
		  // Execute the query which will be the data source of the class
		  
		  Try
		    DBaseRS=DBaseID.SelectSQL(SqlQuerySource,SqlParams())
		  Catch error As DatabaseException
		    MessageBox("Bad Sql resquest: " + error.Message)
		    return false
		  End Try
		  
		  
		  Try
		    DBaseRS.MoveToFirstRow
		  Catch error As DatabaseException
		    MessageBox("DB Error: " + error.Message)
		    return false
		  End Try
		  
		  Fields_Load()
		  
		  return true
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		BLOCAGE As String = "N"
	#tag EndProperty

	#tag Property, Flags = &h0
		DBaseID As SQLiteDatabase
	#tag EndProperty

	#tag Property, Flags = &h0
		DBaseRS As RowSet
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// [Update-XojoCode-NextLines]
		#tag EndNote
		Field_BirthDay As DateTime
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// [Update-XojoCode-NextLines]
		#tag EndNote
		Field_CreditAmount As double = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// [Update-XojoCode-NextLines]
		#tag EndNote
		Field_FirstName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// [Update-XojoCode-NextLines]
		#tag EndNote
		Field_LastName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// Name of the key identifying the record in the table
			// [Update-XojoCode-NextLines]
		#tag EndNote
		KeyTableName As string = "KEY_ID"
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// [Update-XojoCode-NextLines]
		#tag EndNote
		KeyTableValue As Int64 = -1
	#tag EndProperty

	#tag Property, Flags = &h0
		SqlParams() As Variant
	#tag EndProperty

	#tag Property, Flags = &h0
		SqlQuerySource As String
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// Name of the table
			// [Update-XojoCode-NextLines]
		#tag EndNote
		TableName As string = "CUSTOMERS"
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="BLOCAGE"
			Visible=false
			Group="Behavior"
			InitialValue="N"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
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
			Name="Field_FirstName"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
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
		#tag ViewProperty
			Name="KeyTableValue"
			Visible=false
			Group="Behavior"
			InitialValue="-1"
			Type="Int64"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Field_LastName"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="KeyTableName"
			Visible=false
			Group="Behavior"
			InitialValue="REFNUM"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TableName"
			Visible=false
			Group="Behavior"
			InitialValue="CUSTOMERS"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Field_CreditAmount"
			Visible=false
			Group="Behavior"
			InitialValue="0"
			Type="double"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="SqlQuerySource"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass

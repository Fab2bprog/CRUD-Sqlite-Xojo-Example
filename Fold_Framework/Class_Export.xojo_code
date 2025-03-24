#tag Class
Protected Class Class_Export
	#tag Method, Flags = &h21
		Private Function ColExcel(numcol as Integer) As string
		  dim alphabet as String
		  dim lettre As string
		  alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		  lettre=alphabet.Middle(numcol,1)
		  
		  return lettre
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Export_CSV() As Boolean
		  Dim stream as TextOutputStream
		  Dim col As  Integer
		  Dim StrLine as String
		  Dim StrField As String
		  
		  if File_Export=NIL then
		    return false
		  end
		  
		  // Abort if a file with the same name exists
		  if File_Export.Exists then
		    return false
		  end
		  
		  stream=TextOutputStream.Create(File_Export)
		  
		  DBaseRS.MoveToFirstRow
		  
		  StrLine=""
		  for col = 0 to DBaseRS.ColumnCount-1
		    
		    StrField= DBaseRS.ColumnAt(col).Name.ReplaceAll(";","[?]")
		    StrField= StrField.ReplaceAll(EndOfLine,"[?]")
		    StrLine =StrLine + StrField+";"
		    
		  next col
		  Stream.WriteLine(StrLine)
		  
		  
		  While not DBaseRS.AfterLastRow
		    StrLine=""
		    for col  = 0 to DBaseRS.ColumnCount-1
		      
		      StrField = DBaseRS.ColumnAt(col).StringValue.ReplaceAll(";","[?]")
		      StrField = StrField.ReplaceAll( EndOfLine,"[?]")
		      StrLine        = StrLine + StrField+";"
		      
		    next col
		    Stream.WriteLine(StrLine)
		    DBaseRS.MoveToNextRow
		  Wend
		  
		  
		  Stream.Close
		  
		  
		  Return True
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Export_Excel() As Boolean
		  
		  // THIS CODE CAN ONLY BE EXECUTED ON PC WINDOWS PLATFORM
		  // BIG WARNING : You will need to copy the MSOfficeAutomation plugin (located in the Extras folder of the installation) to the Plugins folder before you can use this class or compile this program with this code uncommented, restart xojo.
		  
		  // [Update-XojoCode-NextLines]
		  // START of the xojo code to uncomment to unlock Excel Export, PLEASE READ WARNING
		  
		  '#If TargetWindows Then
		  'Try
		  'Dim myexcel As New ExcelApplication
		  'Dim mybook As ExcelWorkbook
		  'Dim lig As Integer = 0
		  'Dim col As  Integer
		  '
		  'myexcel.Visible = True
		  'mybook = myexcel.Workbooks.Add
		  'myexcel.ActiveSheet.Name = "Export_page_client"
		  '
		  'DBaseRS.MoveToFirstRow
		  '
		  'for col = 0 to DBaseRS.ColumnCount-1
		  'myexcel.Range(ColExcel(col) + Str(1), ColExcel(col) + Str(1)).Value = DBaseRS.ColumnAt(col).Name
		  'next col
		  '
		  'While not DBaseRS.AfterLastRow
		  '
		  'lig=lig+1
		  '
		  'for col  = 0 to DBaseRS.ColumnCount-1
		  'myexcel.Range(ColExcel(col) + Str(lig + 1), ColExcel(col) + Str(lig + 1)).Value =DBaseRS.ColumnAt(col).StringValue
		  'next 
		  '
		  'DBaseRS.MoveToNextRow
		  '
		  'Wend
		  '
		  'return true
		  '
		  'Catch error As OLEException
		  'MessageBox "To export to Excel format your PC must be running a Windows operating system with Microsoft Excel installed !"+EndOfLine+error.message
		  'return false
		  'End Try
		  '
		  'return true
		  '
		  '#Endif
		  
		  // END of the xojo code to uncomment to unlock Excel Export
		  
		  MessageBox "To export to Excel format, your PC must be running a Windows operating system with Microsoft Excel installed, otherwise CSV format will be selected. "+EndOfLine
		  
		  return false
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Run_SqlQuerySource() As Boolean
		  
		  Try
		    DBaseRS=DBaseID.SelectSQL(SqlQuerySource)
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
		  
		  return true
		  
		  
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Select_File(NameFormat as string = "texte", ExtFichier as String = "txt") As Boolean
		  dim dlog as SaveFileDialog
		  
		  
		  // Creates a file creation dialog box
		  dlog = New SaveFileDialog
		  dlog.Title="Exporting the current page to the format "+ NameFormat
		  dlog.SuggestedFileName = "ExportedFile."+ExtFichier
		  File_Export = dlog.ShowModal
		  
		  // In case the user cancels the choice of a file
		  if File_Export = NIL then
		    return false
		  end
		  
		  if File_Export.Exists then
		    File_Export.Remove
		  end
		  
		  return true
		  
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		DBaseID As SQLiteDatabase
	#tag EndProperty

	#tag Property, Flags = &h0
		DBaseRS As RowSet
	#tag EndProperty

	#tag Property, Flags = &h0
		File_Export As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h0
		SqlQuerySource As String
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="SqlQuerySource"
			Visible=false
			Group="Behavior"
			InitialValue=""
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

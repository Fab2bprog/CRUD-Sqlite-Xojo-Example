#tag Class
Protected Class App
Inherits DesktopApplication
	#tag Method, Flags = &h0
		Sub CloseDatabaseFile()
		  Dim message as string
		  
		  if not (MainDB_State="CLOSE") then
		    message =  "Closing the database  : "+EndOfLine+MainDB.DatabaseFile.Name
		    MainDB.Close
		    MessageBox message
		    MainDB=new SQLiteDatabase
		    MainDB_State="CLOSE"
		  else
		    MessageBox  "No open database !"
		  end if
		  
		  Exception err as NilObjectException
		    MainDB_State="CLOSE"
		    MessageBox "Error : No open database !"
		    
		    
		    
		    
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CreateDatabaseFile() As Boolean
		  Dim MyXojoDBDict as new Class_XojoDBDict
		  
		  
		  Try
		    // Create the database file
		    MainDB.CreateDatabase 
		    MyXojoDBDict.DBaseID=App.MainDB
		    MyXojoDBDict.Initialise_Base
		    return App.OpenDatabaseFile
		  Catch error As IOException
		    MessageBox("The database file could not be created: " + error.Message)
		    return false
		  End Try
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function OpenDatabaseFile() As Boolean
		  
		  Try
		    MainDB.Connect 
		    MainDB_State="OPEN"
		    return true
		  Catch error As DatabaseException
		    MessageBox("Error connecting to the database: " +EndOfLine+ error.Message)
		    return false
		  End Try
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function RunSqlQuery(SqlOrder as String) As Boolean
		  Try
		    MainDB.ExecuteSQL(SqlOrder)
		    Return true
		  Catch error As DatabaseException
		    MessageBox("DB Error: " + error.Message)
		    Return false
		  End Try
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function SuperRound(nombre as double, nbrdec as Integer) As Double
		  // Procedure for rounding numbers in a better way than native xojo instructions
		  
		  Dim i as Int64
		  Dim multipli as int64
		  
		  multipli=1
		  
		  for i=1 to nbrdec
		    multipli = 10 * multipli
		  next i
		  
		  nombre = nombre * multipli
		  nombre = round(nombre)
		  nombre = nombre/multipli
		  
		  return nombre
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Valide_TxtNumKeydown(byref myNumFieldTxt as DesktopTextField, Key as string, IsAnInteger as Boolean = false) As Boolean
		  // The purpose of this procedure is to manage the entry of numbers according to certain criteria.
		  // This procedure must be called from a text field during the KeyDown event
		  
		  if IsNumeric( Key ) then
		    
		    // In case you try to write a number before the minus sign the cursor is placed at the end of the TargetField
		    if myNumFieldTxt.SelectionStart=0 and myNumFieldTxt.SelectionLength=0 and myNumFieldTxt.Text.Left(1)="-"  then
		      myNumFieldTxt.SelectionStart=myNumFieldTxt.Text.Length
		    end if
		    
		    // All numbers must be entered
		    return false
		    
		  end if
		  
		  // Handling the case where the person wants to enter decimals in an integer field
		  if Key = SymbDecimal and IsAnInteger then
		    return true
		  end if
		  
		  //Managing cases where the person types the decimal symbol 
		  if ( Key = SymbDecimal ) and myNumFieldTxt.Text.Contains(SymbDecimal)=false  then
		    
		    if myNumFieldTxt.Text.Length = 0 or myNumFieldTxt.SelectionLength=myNumFieldTxt.Text.Length then
		      myNumFieldTxt.Text="0"
		      myNumFieldTxt.SelectionStart=myNumFieldTxt.Text.Length
		    end if
		    
		    if myNumFieldTxt.Text= "-" then
		      myNumFieldTxt.Text="-0"
		      myNumFieldTxt.SelectionStart=myNumFieldTxt.Text.Length
		    end if
		    
		    return false
		    
		  end if
		  
		  // Allows certain keys like backspace and delete
		  if ASC( Key ) < 32 or  ASC( Key ) =127 then
		    return false
		  end if
		  
		  
		  //Management of the minus sign
		  if ASC( Key ) = 45 then
		    
		    if myNumFieldTxt.Text.Length >0 then
		      
		      if myNumFieldTxt.Text.Left(1)="-"  then
		        // Transformation of a negative number into a positive number
		        myNumFieldTxt.Text=myNumFieldTxt.Text.Right(myNumFieldTxt.Text.Length - 1 )
		      else
		        // Transformation of a positive number into a negative number
		        myNumFieldTxt.Text="-"+myNumFieldTxt.Text
		      end if
		      
		    else
		      myNumFieldTxt.text= "-"
		    end if
		    
		    // We place the cursor at the end
		    myNumFieldTxt.SelectionStart=myNumFieldTxt.Text.Length
		    
		  end if
		  
		  
		  
		  // Everything that has not been permitted is forbidden.
		  return true
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Valide_TxtNumLostFocus(byref myNumFieldTxt as DesktopTextField, IsAnInteger as Boolean = false, Precision as integer = 0, ValMin as Double = -9999999999999, ValMax as Double = 9999999999999, ValDefaut as double = 0) As Boolean
		  // The purpose of this procedure is to manage the entry of numbers according to certain criteria.
		  // This procedure must be called from a text field during the LostFocus event
		  
		  Dim MyNbr as Double
		  
		  // Cleaning up spaces and tabs that could result from an unfortunate copy and paste
		  myNumFieldTxt.Text=myNumFieldTxt.Text.ReplaceAll(" ","")
		  myNumFieldTxt.Text=myNumFieldTxt.Text.ReplaceAll(chr(9),"")
		  
		  
		  // In case the field is empty and a default value has been defined
		  if myNumFieldTxt.Text.Length=0  then
		    myNumFieldTxt.Text=ValDefaut.ToString
		    return true
		  end if
		  
		  // If the field is an integer and a decimal value has been entered, it is rounded down to the nearest integer.
		  if IsAnInteger then
		    myNumFieldTxt.Text=floor(myNumFieldTxt.Text.CDbl).ToString
		    return true
		  end if
		  
		  MyNbr = myNumFieldTxt.Text.CDbl
		  
		  // If the user enters more decimals than necessary, the value is rounded.
		  if Precision > 0 then
		    MyNbr = SuperRound(myNumFieldTxt.Text.CDbl,precision)
		  end if
		  
		  // Case where the user has exceeded the minimum or maximum values
		  if MyNbr<ValMin then
		    MessageBox "The value you entered ("+MyNbr.ToString+") is less than the minimum allowed value ("+ValMin.ToString+"), "+EndOfLine+_
		    "the program assigned the minimum value to the field"
		    myNumFieldTxt.Text = ValMin.ToString
		    return false
		  end if
		  
		  if MyNbr>ValMax then
		    MessageBox "The value you entered ("+MyNbr.ToString+") is greater than the maximum allowed value ("+ValMin.ToString+"), "+EndOfLine+_
		    "the program assigned the maximum value to the field"
		    myNumFieldTxt.Text = ValMin.ToString
		    return false
		  end if
		  
		  
		  //Everything that has not been forbidden is permitted
		  myNumFieldTxt.Text=MyNbr.ToString
		  return true
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod


	#tag Note, Name = Licence MIT
		MIT License
		
		Copyright (c) 2025 Fabrice Garcia, 20290 Borgo, Corsica, France, Europe.
		
		Permission is hereby granted, free of charge, to any person obtaining a copy
		of this software and associated documentation files (the "Software"), to deal
		in the Software without restriction, including without limitation the rights
		to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
		copies of the Software, and to permit persons to whom the Software is
		furnished to do so, subject to the following conditions:
		
		The above copyright notice and this permission notice shall be included in all
		copies or substantial portions of the Software.
		
		THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
		IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
		FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
		AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
		LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
		OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
		SOFTWARE.
		
		
	#tag EndNote

	#tag Note, Name = ReadMe: About missing images
		When you first open the project, Xojo may not find the images.
		This is a common problem when a Xojo project is moved to another computer.
		The images are all located in the pictures folder : Don't panic, none are missing.
		In the alert window asking you to locate the missing images, check all the images and click the Resolve button, go to the pictures folder ( inside the project folder ) and designate them.
		The images used were all generated by the software author using AI and are royalty-free. You are free to use them in other projects as you wish. They are also covered by the MIT license.
		
		
	#tag EndNote

	#tag Note, Name = ReadMe: Adapting to your needs
		I wrote in the comments a tag // [Update-XojoCode-NextLines]
		When you see this tag in comments, it means that the program probably needs to be modified in the following lines to meet your needs.
		If you see this tag in the declaration of a variable, it means that you will probably need to modify the name of the variable or its default value to adapt the program to your needs.
		Please note: This does not mean that depending on what you want to do, you will only need to modify these lines.
		This is simply an aid to help you locate important points that most likely need to be modified to suit your needs. 
		I do not claim to be exhaustive on this subject.
		
	#tag EndNote

	#tag Note, Name = ReadMe: Excel data export
		------------------------------------------------------
		* Procedure to unlock data export to Microsoft Excel *
		------------------------------------------------------
		
		1- You will need to copy the MSOfficeAutomation plugin (located in the Extras folder of the xojo installation) to the Plugins folder before you can use this functionality.
		2- Exit Xojo and relaunch the application
		3- You must have Microsoft Excel installed on your computer, only works on Windows PC.
		4- Open Class Class_Export, open methods Export_Excel : read the comments
		5- Inside Class_Export.Export_Excel : Uncomment the area defined by the comments (From: #If Target Windows ... until #Endif)
		
		
		
		
		
	#tag EndNote

	#tag Note, Name = ReadMe: Project structure
		The project includes 3 windows :
		- A start window ( Win_Welcome ) that includes the menu for opening, creating, and closing databases and a number of indications.
		- A data list window ( Win_CustomerList ) that allows you to Create, Read, Modify, Delete records from the SQLite database.
		- A customer data entry form (Win_CustomerForm)
		
		________________________________
		
		At the Object class level :
		
		1 - Class_Customer : 
		This class allows you to create an abstraction layer to manage data from the SQLite database.
		This class allows you to create, add, modify, delete, read records from the database.
		The advantages of using a class and not typing instructions directly into the project windows are numerous.
		This greatly facilitates project maintenance and avoids code duplication. 
		Furthermore, if you ever want to change the database type and the database operations need to be handled differently, you'll only have to modify this class. 
		There's no need to travel to all the project windows to modify code, which is tiring.
		In addition, this makes the code more independent of Xojo instructions.
		In the event that these instructions change or their uses change, your program is not to be thrown away. 
		Using classes ensures a certain longevity of the project.
		
		2 - Class_XojoDBDict :
		The purpose of this class is to initialize the SQLite database.
		This class creates the blank tables necessary for the program to function.
		
		3- Class_MenuTable and Class_MenuWelcome
		Manages the program menus.
		
		4- Class_Export
		The purpose of this class is to export the database data in CSV or Microsoft Excel format.
		
		
		
		
		
	#tag EndNote


	#tag Property, Flags = &h0
		MainDB As SQLiteDatabase
	#tag EndProperty

	#tag Property, Flags = &h0
		MainDB_State As String = "CLOSE"
	#tag EndProperty


	#tag Constant, Name = kEditClear, Type = String, Dynamic = False, Default = \"&Delete", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"&Delete"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"&Delete"
	#tag EndConstant

	#tag Constant, Name = kFileQuit, Type = String, Dynamic = False, Default = \"&Quit", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"E&xit"
	#tag EndConstant

	#tag Constant, Name = kFileQuitShortcut, Type = String, Dynamic = False, Default = \"", Scope = Public
		#Tag Instance, Platform = Mac OS, Language = Default, Definition  = \"Cmd+Q"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"Ctrl+Q"
	#tag EndConstant

	#tag Constant, Name = SymbDecimal, Type = String, Dynamic = False, Default = \".", Scope = Public
		#Tag Instance, Platform = Mac OS, Language = Default, Definition  = \"\x2C"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"."
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"."
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=false
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=false
			Group="ID"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=false
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=false
			Group="Position"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=false
			Group="Position"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AllowAutoQuit"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AllowHiDPI"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="BugVersion"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Copyright"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Description"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LastWindowIndex"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="MajorVersion"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="MinorVersion"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="NonReleaseVersion"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="RegionCode"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="StageCode"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Version"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="_CurrentEventTime"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="MainDB_State"
			Visible=false
			Group="Behavior"
			InitialValue="CLOSE"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass

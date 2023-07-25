Private Sub Workbook_Open()
	' Handle opening sql file
	' Regex to grab the portion we need into string specialCodeString 
	Dim 
	specialCodeString.split(vbLf)
	
	Sub PopulateComboBoxWithSplitString()
	    Dim myString As String
	    Dim myArray() As String
	    Dim i As Integer

	    ' Sample string with new line characters
	    myString = "Apple" & vbCrLf & "Banana" & vbCrLf & "Orange" & vbCrLf & "Mango"

	    ' Split the string into an array based on new line characters
	    myArray = Split(myString, vbCrLf)

	    ' Clear the ComboBox before adding new items
	    ComboBox1.Clear

	    ' Loop through the array and add items to the ComboBox
	    For i = LBound(myArray) To UBound(myArray)
		ComboBox1.AddItem myArray(i)
	    Next i
	End Sub

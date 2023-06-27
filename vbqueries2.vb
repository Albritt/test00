Sub ExecuteSQLFile()
    Dim fso As Object
    Dim file As Object
    Dim sqlFilePath As String
    Dim sqlFileContents As String
    Dim sqlStatements() As String
    Dim conn As Object
    Dim cmd As Object
    Dim i As Long
    
    ' Set the path of the SQL file
    sqlFilePath = "C:\Path\to\your\sqlfile.sql"
    
    ' Create a file system object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Read the contents of the SQL file
    sqlFileContents = fso.OpenTextFile(sqlFilePath, 1).ReadAll
    
    ' Split the file into individual SQL statements using "GO" as the delimiter
    sqlStatements = Split(sqlFileContents, "GO", -1, vbTextCompare)
    
    ' Create a new ADO connection
    Set conn = CreateObject("ADODB.Connection")
    
    ' Set connection properties (e.g., provider, server, database, etc.)
    conn.Provider = "YourProvider"
    conn.ConnectionString = "YourConnectionString"
    conn.Open
    
    ' Create a new ADO command
    Set cmd = CreateObject("ADODB.Command")
    
    ' Set command properties (e.g., connection, timeout, etc.)
    cmd.ActiveConnection = conn
    cmd.CommandTimeout = 300 ' Set the desired timeout value
    
    ' Loop through each SQL statement and execute it
    For i = 0 To UBound(sqlStatements)
        cmd.CommandText = sqlStatements(i)
        cmd.Execute
    Next i
    
    ' Close the connection
    conn.Close
    
    MsgBox "SQL file executed successfully."
End Sub

Sub ReplaceGoWithSemicolon()
    Dim filePath As String
    Dim fileContents As String
    Dim regex As Object
    Dim replacedContents As String
    
    ' Set the file path to your SQL script
    filePath = "C:\path\to\your\sql_script.sql"
    
    ' Read the contents of the SQL script file
    With CreateObject("Scripting.FileSystemObject")
        Dim fileStream As Object
        Set fileStream = .OpenTextFile(filePath, 1, False)
        fileContents = fileStream.ReadAll()
        fileStream.Close
    End With
    
    ' Create a regular expression object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Set the pattern to match "GO" statements (case-insensitive)
    regex.Pattern = "(?i)\bGO\b"
    
    ' Replace "GO" with semicolons
    replacedContents = regex.Replace(fileContents, ";")
    
    ' Output the modified SQL script
    Debug.Print replacedContents
    
    ' Now you can use the replacedContents string as needed, such as executing it with ADO
    
End Sub


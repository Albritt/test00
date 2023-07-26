Dim regexPattern As String
Dim inputText As String
Dim result As String

' Your VBA code here (inputText)'

' Define the regex pattern to capture the desired text
regexPattern = "(?<=Put special codes here).*?(?=END)"

' Create a regex object
Dim regex As Object
Set regex = CreateObject("VBScript.RegExp")

' Set regex pattern and ignore case
With regex
    .Global = True
    .MultiLine = True
    .IgnoreCase = True
    .Pattern = regexPattern
End With

' Execute the regex pattern on the input text
Dim matches As Object
Set matches = regex.Execute(inputText)

' If there are matches, extract the captured text
If matches.Count > 0 Then
    result = matches(0).Value
    MsgBox "Captured text: " & result
Else
    MsgBox "No match found."
End If

Sub main() 'multiple send

'Call EditHTMLFile

If Sheets("Name list").Cells(2, "B").Value <> "" Then

Call send_loop

End If

Debug.Print "main"

End Sub

Sub send_loop() 'loop

Dim name As String
Dim email_list As String
Dim First_name As String


Sheets("Name list").Select

table = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To table

    name = Sheets("Name list").Cells(i, 1).Value
    First_name = Sheets("Name list").Cells(i - 1, 1).Value
    Call EditHTMLUsername(name, First_name)
    email_list = Sheets("Name list").Cells(i, 2).Value
    Call sendemail(email_list)
    'Debug.Print name & "111" & First_name
 
Next i


MsgBox "Multiple email send"
End Sub

Sub EditHTMLUsername(name As String, First_name As String)

Dim objFSO As Object
Dim objTextFile As Object
Dim strFilePath As String
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

'file location
H = Sheet2.Cells(2, "C").Value
'image path
newreplace = name 'Sheet1.Cells(2, "A").Value
strFilePath = "" & H & ""

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile(strFilePath, ForReading)

Dim strFileContent As String


If Not objTextFile.AtEndOfStream Then strFileContent = objTextFile.ReadAll 'strFileContent = objTextFile.ReadAll 'If

objTextFile.Close

'start editing the file content here and auto save
strFileContent = Replace(strFileContent, First_name, newreplace)
'Debug.Print strFileContent
Set objTextFile = objFSO.OpenTextFile(strFilePath, ForWriting)
objTextFile.Write strFileContent
objTextFile.Close
Debug.Print "EditHTMLFile"



End Sub


Sub sendemail(email_list As String)

Dim T As String
Dim atA As Attachment
Dim objMessage, objConfig, Fields
E = Sheet2.Cells(2, "J").Value
T = Sheet2.Cells(7, "J").Value
H = Sheet2.Cells(2, "C").Value
'send email
Set objMessage = CreateObject("CDO.Message")
Set objConfig = CreateObject("CDO.Configuration")
Set Fields = objConfig.Fields
With Fields
  .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.office365.com"
  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
  .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
  .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "sales@gmail.com.my" '"linda.septiyana@gmail.com.my" '
  .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Jmp@2023" '"Linda@4644" '
  .Item("http://schemas.microsoft.com/cdo/configuration/sendtls") = True
  .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
  .Update
End With
Set objMessage.Configuration = objConfig

With objMessage
  .Subject = T
  .From = "sales@gmail.com.my"
  .To = email_list
  .CreateMHTMLBody "" & H & ""

End With


objMessage.send

Debug.Print "sendemail"
End Sub

Sub EditHTMLFile()

Dim objFSO As Object
Dim objTextFile As Object
Dim strFilePath As String
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

'file location
H = Sheet2.Cells(2, "C").Value
'image path
newreplace = Sheet2.Cells(2, "E").Value
strFilePath = "" & H & ""

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile(strFilePath, ForReading)

Dim strFileContent As String

On Error GoTo ErrorHandler
If Not objTextFile.AtEndOfStream Then strFileContent = objTextFile.ReadAll 'strFileContent = objTextFile.ReadAll 'If

objTextFile.Close

'start editing the file content here and auto save
strFileContent = Replace(strFileContent, "src=""", "src=""" & newreplace & "")
'Debug.Print strFileContent
Set objTextFile = objFSO.OpenTextFile(strFilePath, ForWriting)
objTextFile.Write strFileContent
objTextFile.Close
Debug.Print "EditHTMLFile"

ErrorHandler:
Debug.Print "Error"
Exit Sub

End Sub

Sub File_Path()
Dim File_Picker As FileDialog
Dim my_path As String
Set File_Picker = Application.FileDialog(msoFileDialogFilePicker)
File_Picker.Title = "Select a File" & FileType
File_Picker.Filters.Clear
File_Picker.Show
If File_Picker.SelectedItems.Count = 1 Then
my_path = File_Picker.SelectedItems(1)
End If

Sheets("Email list").Range("C2").Value = my_path ' fullPath
Sheets("Email list").Range("E2").Value = Left(my_path, InStrRev(my_path, "\")) 'strScriptPath
Sheets("Email list").Range("G2").Value = Right(my_path, Len(my_path) - InStrRev(my_path, "\")) 'strScriptName

Debug.Print "File_Path"
End Sub

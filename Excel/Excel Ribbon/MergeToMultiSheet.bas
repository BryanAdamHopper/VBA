Attribute VB_Name = "MergeToMultiSheet"
Sub MergeToMultiSheet()
Dim MyPath As String
Dim FilesInPath As String
Dim MyFiles() As String
Dim SourceRcount As Long
Dim Fnum As Long
Dim mybook As Workbook
Dim basebook As Workbook

'Fill in the path\folder where the files are
'on your machine
'MyPath = "c:\Data"
    Set myFile = Application.FileDialog(msoFileDialogFolderPicker)
    With myFile
        .Title = "Choose File"
        .AllowMultiSelect = False
    If .Show <> -1 Then
    Exit Sub
    End If
    FileSelected = .SelectedItems(1)
    End With
MyPath = FileSelected

'Add a slash at the end if the user forget it
If Right(MyPath, 1) <> "\" Then
MyPath = MyPath & "\"
End If

'If there are no Excel files in the folder exit the sub
'FilesInPath = Dir(MyPath & "*.csv")
myExt = Application.InputBox("Enter an extention, with no dot.")
FilesInPath = Dir(MyPath & "*." & myExt)

If FilesInPath = "" Then
MsgBox "No files found"
Exit Sub
End If

On Error GoTo CleanUp

Application.ScreenUpdating = False
Set basebook = ThisWorkbook

'Fill the array(myFiles)with the list of Excel files in the folder
Fnum = 0
Do While FilesInPath <> ""
Fnum = Fnum + 1
ReDim Preserve MyFiles(1 To Fnum)
MyFiles(Fnum) = FilesInPath
FilesInPath = Dir()
Loop

'Loop through all files in the array(myFiles)
If Fnum > 0 Then
For Fnum = LBound(MyFiles) To UBound(MyFiles)
Set mybook = Workbooks.Open(MyPath & MyFiles(Fnum))
mybook.Worksheets(1).Copy after:= _
basebook.Sheets(basebook.Sheets.Count)

On Error Resume Next
'ActiveSheet.Name = mybook.Name 'remove this link to remove ext being added to tab name.

On Error GoTo 0

' You can use this if you want to copy only the values
' With ActiveSheet.UsedRange
' .Value = .Value
' End With

mybook.Close savechanges:=False
Next Fnum
End If
CleanUp:
Application.ScreenUpdating = True
End Sub




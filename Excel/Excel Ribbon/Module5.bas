Attribute VB_Name = "Module5"
Sub DeleteAllPics()

Dim Pic As Object

For Each Pic In ActiveSheet.Pictures
    'Pic.Delete
    'MsgBox (Pic.Name)
    MsgBox (ActiveSheet.cell)
Next Pic

End Sub

Sub SaveAllPics()

Dim oTxt As Object
'Dim WS As Worksheet
'Set WS = ActiveWorksheet

Set FOLDER = Application.FileDialog(msoFileDialogFolderPicker)
FOLDER.AllowMultiSelect = False


 For Each cell In ActiveWorksheet.Range("A1:A" & ActiveWorksheet.UsedRange.Rows.Count)
    saveText = cell.Text
    Open FOLDER & saveText & ".jpg" For Output As #1
    Print #1, cell.Offset(0, 1).Text
    Close #1
 Next cell

End Sub


Sub ExportPictures()

Dim WS As Worksheet
Dim Pic As Object

    '## Open file dialog to choose a destination folder
    Set FOLDER = Application.FileDialog(msoFileDialogFolderPicker)
    FOLDER.AllowMultiSelect = False
    FOLDER.Show

    '## loop through all sheets and all pictures
    For Each WS In ThisWorkbook.Sheets
    For Each Pic In WS.Shapes

        '## create a chart with same dimensions as current picture
        '## subtract 0.5px from chart dimensions to avoid a strange border
        Set CH = WS.ChartObjects.Add(1, 1, Pic.Height, Pic.Width)

        '## save & temporarly disable the picture border
        Pic.Select
        PICBORDER = Selection.Border.LineStyle
        Selection.Border.LineStyle = 0

        '## copy the picture into chart. Only a chart could be exported
        Pic.Copy
        CH.Chart.ChartArea.Select
        CH.Chart.Paste

        '## re-enable the old picture border
        Pic.Select
        Selection.Border.LineStyle = PICBORDER

        '## export the chart as JPG. Change JPG to PNG if desired
        CH.Chart.Export Filename:=FOLDER.SelectedItems(1) & "\" & Pic.Name & ".jpg", FilterName:="JPG"

        '## delete chart to clean up our work
        CH.Cut

    Next Pic
    Next WS
End Sub




Option Explicit

Sub ExportMyPicture()

     Dim MyChart As String, MyPicture As String
     Dim PicWidth As Long, PicHeight As Long

     Application.ScreenUpdating = False
     On Error GoTo Finish

     MyPicture = Selection.Name
     With Selection
           PicHeight = .ShapeRange.Height
           PicWidth = .ShapeRange.Width
     End With

     Charts.Add
     ActiveChart.Location Where:=xlLocationAsObject, Name:="Sheet1"
     Selection.Border.LineStyle = 0
     MyChart = Selection.Name & " " & Split(ActiveChart.Name, " ")(2)

     With ActiveSheet
           With .Shapes(MyChart)
                 .Width = PicWidth
                 .Height = PicHeight
           End With

           .Shapes(MyPicture).Copy

           With ActiveChart
                 .ChartArea.Select
                 .Paste
           End With

           .ChartObjects(1).Chart.Export Filename:="MyPic.jpg", FilterName:="jpg"
           .Shapes(MyChart).Cut
     End With

     Application.ScreenUpdating = True
     Exit Sub

Finish:
     MsgBox "You must select a picture"
End Sub

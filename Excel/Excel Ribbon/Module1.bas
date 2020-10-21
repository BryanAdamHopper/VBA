Attribute VB_Name = "Module1"

'Defines column to search
Private Const SEARCH_COLUMN = "A:A"

Sub confMacro()
ans = MsgBox("Are you sure?", vbYesNo, "Confirmation")
If ans = vbYes Then RemoveAlphas 'if not Yes as reply, naturally exit sub
End Sub

Sub BackupGen()

'NOTE: To call this function use: "Call BackupGen"

'Get active worksheet name.
ActiveWSNameOrig = ActiveSheet.Name 'Added due to var being altered but still needing the original name.
ActiveWSName = ActiveSheet.Name

'Create new worksheet as "BAK-<WorksheetName>-<TimeStamp>-<"RN">-<RandomNumber>".
'Name cannot exceed 31 characters.

WSNameLen = Len(ActiveWSName)
WSNameDefaultLen = Len("BAKxxxxXX")
MaxLen = (30 - WSNameDefaultLen)

    'Check to see if current worksheet name needs to be truncated so new name does not exceed 31 characters.
    If WSNameLen >= (MaxLen - 1) Then
        ActiveWSName = Mid(ActiveWSName, 1, MaxLen)
    End If

Dim WS As Worksheet
Set WS = Sheets.Add(after:=Sheets(Worksheets.Count))
Time_Stamp = Format(Now(), "yyyy-MM-dd hh:mm:ss")
Time_Stamp = Mid(Time_Stamp, 12, 2) + Mid(Time_Stamp, 15, 2) + CStr(Int(89 * Rnd) + 10)
        LenTestTemp = Len("BAK" + ActiveWSName + Time_Stamp)
WS.Name = "BAK" + ActiveWSName + Time_Stamp
Sheets(ActiveWSNameOrig).Activate

'Copy active worksheet to backup worksheet.
ActiveSheet.Cells.Copy
Worksheets(WS.Name).Select
Cells.Select
Selection.PasteSpecial Paste:=xlPasteAllExceptBorders, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Sheets(ActiveWSNameOrig).Activate

End Sub

Sub CheckAlpha()

    Dim Sure As Integer

    Sure = MsgBox("Are you sure?", vbOKCancel)
    If Sure = 1 Then Call RemoveAlphas

End Sub

Sub Gender()
'
' Gender Macro
'

'Created backup before proceeding.
Call BackupGen

    Selection.Replace What:="Female", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="F", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Male", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="M", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

Sub BoldFreeze()
'
' BoldFreeze Macro
'
' Keyboard Shortcut: Ctrl+q
'
    Rows("1:1").Select
    Selection.Font.Bold = True
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Sub Autofilter()
'
' Autofilter Macro
'
' Keyboard Shortcut: Ctrl+w
'
    Cells.Select
    Selection.Autofilter
    Selection.Rows.AutoFit
    Selection.Columns.AutoFit
End Sub

Sub ColumnHeaders()
Set shtJT = ActiveWorkbook.ActiveSheet


    ActiveWorkbook.ActiveSheet.Select
    Rows("1:1").Select
    Selection.Copy
    Sheets.Add after:=Sheets(Sheets.Count)
    Range("B1").Select 'Just select a single cell, not the whole column
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
    True, Transpose:=True
    Selection.Columns.AutoFit
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "A"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "B"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "C"
    Range("A1:A3").Select
    Selection.AutoFill Destination:=Range("A1:A94"), Type:=xlFillDefault
    Range("A1:A94").Select
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""Column "",RC[-2],"", "",RC[-1])"
    Range("C1").Select
    Selection.Copy
    Range("C1:C94").Select
    ActiveSheet.Paste
    Range("C:C").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Selection.Columns.AutoFit

End Sub

Sub Replace()

Dim choiceVal As Integer
Dim MsgBoxText As String

    MsgBoxText = _
                "Enter a value (1-4):" & Chr(13) & Chr(13) & _
                "1 - Text to Char (X to 000)" & Chr(13) & _
                "2 - Char to Text (000 to X)" & Chr(13) & _
                "3 - Char to Char (000 to 000)" & Chr(13) & _
                "4 - Text to Text (X to X)"
    choiceVal = Application.InputBox(prompt:=MsgBoxText, Title:="Special Replace", Type:=1)
     
If choiceVal = 0 Then
    Exit Sub
    
ElseIf choiceVal = 1 Then
    Selection.Replace _
        InputBox(prompt:="Text to replace", Title:="Old Text", Default:="X"), _
        Chr(InputBox(prompt:="Replace text with this char", Title:="New Char Code", Default:="000"))
        
ElseIf choiceVal = 2 Then
    Selection.Replace _
        Chr(InputBox(prompt:="Char to replace", Title:="Old Char Code", Default:="000")), _
        InputBox(prompt:="Replace char with this text", Title:="New Text", Default:="X")
          
ElseIf choiceVal = 3 Then
    Selection.Replace _
        Chr(InputBox(prompt:="Char to replace", Title:="Old Char Code", Default:="000")), _
        Chr(InputBox(prompt:="Replace char with this char", Title:="New Char Code", Default:="000"))

ElseIf choiceVal = 4 Then
    Selection.Replace _
        InputBox(prompt:="Text to replace", Title:="Old Text", Default:="000"), _
        InputBox(prompt:="Replace text with this text", Title:="New Text", Default:="000")

Else
    MsgBox ("Not a valid option.")
End If

End Sub

Sub Proper_Case()

    For Each x In Selection
        x.Value = Application.Proper(x.Value)
    Next
    
End Sub

Sub RemoveAlphas()

'Created backup before proceeding.
Call BackupGen

Dim intI As Integer
Dim rngR As Range, rngRR As Range
Dim strNotNum As String, strTemp As String
Set rngRR = Selection.SpecialCells(xlCellTypeConstants, _
xlTextValues)
For Each rngR In rngRR
strTemp = ""
For intI = 1 To Len(rngR.Value)
If Mid(rngR.Value, intI, 1) Like "[0-9.]" Then
strNotNum = Mid(rngR.Value, intI, 1)
Else: strNotNum = ""
End If
strTemp = strTemp & strNotNum
Next intI
rngR.Value = strTemp
Next rngR

Selection.Replace What:=".", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
                
End Sub

Sub RemoveHyperlinks()
    ActiveSheet.Hyperlinks.Delete
End Sub

Sub RemoveHidden()
Call BackupGen
Dim oneCell, rngRR As Range
Set rngRR = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)
For Each oneCell In rngRR
With oneCell
    .Value = Evaluate("IF(ISTEXT(" & .Address & "),TRIM(SUBSTITUTE(" & .Address & ",CHAR(160),"" "")),REPT(" & .Address & ",1))")
End With
Next oneCell
End Sub

Sub Calib10()
'
' Calib10 Macro
'

    Cells.Select
    Range("F15").Activate
    With Selection.Font
        .Name = "Calibri"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    'Selection.Font.Size = 10
    'Selection.Font.Size = 9
    'Selection.Font.Size = 8
    Cells.EntireColumn.AutoFit
    Range("E9").Select
End Sub

Sub HeaderChange1()
'
' HeaderChange1 Macro
'
    Rows("1:1").Select
    
    Selection.Replace What:="First Name", Replacement:="FirstName", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Selection.Replace What:="last Name", Replacement:="LastName", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Selection.Replace What:="Email", Replacement:="EmailName", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Selection.Replace What:="E-Mail", Replacement:="EmailName", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Selection.Replace What:="Mobile Phone", Replacement:="Cellphone", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Selection.Replace What:="Birthday", Replacement:="BirthDate", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
End Sub

Sub YN2Bit()
'
' YN2Bit Macro
'

'Backup sheet before proceeding
Call BackupGen

    Selection.Replace What:="Yes", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="No", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Y", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="N", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

Sub PasswordBreaker()
    'Breaks worksheet password protection.
    Dim i As Integer, j As Integer, k As Integer
    Dim l As Integer, m As Integer, n As Integer
    Dim i1 As Integer, i2 As Integer, i3 As Integer
    Dim i4 As Integer, i5 As Integer, i6 As Integer
    On Error Resume Next
    For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
    For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
    For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
    For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
    ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _
        Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
        Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
    If ActiveSheet.ProtectContents = False Then
        MsgBox "One usable password is " & Chr(i) & Chr(j) & _
            Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
            Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
         Exit Sub
    End If
    Next: Next: Next: Next: Next: Next
    Next: Next: Next: Next: Next: Next
End Sub


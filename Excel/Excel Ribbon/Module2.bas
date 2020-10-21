Attribute VB_Name = "Module2"
Sub RunMacro1()
'
' Run Macro Command
'

frmRunMacro.Show

End Sub

Sub Column_Width()
'
' Column_Width Macro
' Auto space column width.
'
' Keyboard Shortcut: Ctrl+w
'
    Columns("A:A").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Columns("A:R").EntireColumn.AutoFit
End Sub

Sub TTC()
    
    Dim FirstCol%, LastCol%
    FirstCol = Selection(1, 1).Column
    LastCol = Range([A1], Selection).Columns.Count
    
    For x = FirstCol To LastCol
     
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Columns(x).TextToColumns DataType:=xlDelimited, _
        ConsecutiveDelimiter:=False, Space:=False
        
    Next x
    
End Sub

Sub RemoveAlphas2()

'   Takes column with alphanumerics and outputs column of numerics only.
'   Assumes alpha and numeric characters are randomly distributed within the cell.

    Dim rngData As Range
    Dim aryData() As Variant
    Dim strNumerics As String
    Dim i As Long, j As Integer

    Set rngData = Intersect(ActiveSheet.UsedRange, Columns(SEARCH_COLUMN))
    aryData = rngData

'   Extract numerics from each cell
    For i = LBound(aryData) To UBound(aryData)
        strNumerics = ""
        For j = 1 To Len(aryData(i, 1))
            If IsNumeric(Mid(aryData(i, 1), j, 1)) Then
                strNumerics = strNumerics & Mid(aryData(i, 1), j, 1)
            End If
        Next j
    
    aryData(i, 1) = strNumerics
    Next i

'   Output to new column
    rngData.EntireColumn.Insert
    rngData.Offset(0, -1) = aryData

End Sub

Sub txtColorBtn()


' "=PERSONAL.XLSB!txtColor(" & #VARIABLE# & ")"
End Sub

Sub bgColorBtn()


' "=PERSONAL.XLSB!BGColor(" & #VARIABLE# & ")"
End Sub

Function txtColor(rng As Range)
    txtColor = rng.Font.ColorIndex
End Function

Function BGColor(rng As Range)
    BGColor = rng.Interior.ColorIndex
End Function

Function CellColorIndex(rng As Range, Optional IsBGColor As Boolean = False)
    If IsBGColor = True Then
       'Get Background Color
       CellColorIndex = rng.Interior.ColorIndex
    Else
       'Get Font Color
       CellColorIndex = rng.Font.ColorIndex
    End If
End Function

Sub TestMessage(message As String)

'message = "ActiveSheet."

MsgBox (message)

Return

End Sub

Sub ValidCCNum()
'
' Validate CC Numbers Macro
'
    
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "CCisValid"
    Range("A2").Select
    'ActiveCell.FormulaR1C1 = "Personal.xlsb!ValidateCCNumber(B2)"
    ActiveCell.Value = "=Personal.xlsb!ValidateCCNumber(B2)"
End Sub

Function getCCtype(ccNum As String)
a = Trim(ccNum)

'a = Left(a, 6)
'VISA: 4XXXXX
'Mastercard: 51XXXX - 55XXXX
'Discover: 6011XX,644XXX,65XXXX
'Amex: 34XXXX,37XXXX

a = Left(a, 1)

Select Case a
    Case 3
        b = "Amex"
    Case 4
        b = "Visa"
    Case 5
        b = "Mastercard"
    Case 6
        b = "Discover"
    Case Else
        b = "Other"
End Select

getCCtype = b

End Function

Function ValidateCCNumber(ccNum As String)
'Function ValidateCCNumber(ccNum1 As String, ccNum2 As String)
'ccNum = Trim(ccNum1) + Trim(ccNum2)


'Fix for 15 character CC numbers such as Amex cards
ccLength = Len(ccNum)
If ccLength = 15 Then
    ccNum = "0" & ccNum
End If


ccIndexNum = 15
If Mid(ccNum, ccIndexNum, 1) * 2 > 9 Then
    tempValue = 1 + ((Mid(ccNum, ccIndexNum, 1) * 2) Mod 10)
Else
    tempValue = Mid(ccNum, ccIndexNum, 1) * 2
End If

ccIndexNum = ccIndexNum - 2
If Mid(ccNum, ccIndexNum, 1) * 2 > 9 Then
    tempValue = tempValue + (1 + ((Mid(ccNum, ccIndexNum, 1) * 2) Mod 10))
Else
    tempValue = tempValue + (Mid(ccNum, ccIndexNum, 1) * 2)
End If

ccIndexNum = ccIndexNum - 2
If Mid(ccNum, ccIndexNum, 1) * 2 > 9 Then
    tempValue = tempValue + (1 + ((Mid(ccNum, ccIndexNum, 1) * 2) Mod 10))
Else
    tempValue = tempValue + (Mid(ccNum, ccIndexNum, 1) * 2)
End If

ccIndexNum = ccIndexNum - 2
If Mid(ccNum, ccIndexNum, 1) * 2 > 9 Then
    tempValue = tempValue + (1 + ((Mid(ccNum, ccIndexNum, 1) * 2) Mod 10))
Else
    tempValue = tempValue + (Mid(ccNum, ccIndexNum, 1) * 2)
End If

ccIndexNum = ccIndexNum - 2
If Mid(ccNum, ccIndexNum, 1) * 2 > 9 Then
    tempValue = tempValue + (1 + ((Mid(ccNum, ccIndexNum, 1) * 2) Mod 10))
Else
    tempValue = tempValue + (Mid(ccNum, ccIndexNum, 1) * 2)
End If

ccIndexNum = ccIndexNum - 2
If Mid(ccNum, ccIndexNum, 1) * 2 > 9 Then
    tempValue = tempValue + (1 + ((Mid(ccNum, ccIndexNum, 1) * 2) Mod 10))
Else
    tempValue = tempValue + (Mid(ccNum, ccIndexNum, 1) * 2)
End If

ccIndexNum = ccIndexNum - 2
If Mid(ccNum, ccIndexNum, 1) * 2 > 9 Then
    tempValue = tempValue + (1 + ((Mid(ccNum, ccIndexNum, 1) * 2) Mod 10))
Else
    tempValue = tempValue + (Mid(ccNum, ccIndexNum, 1) * 2)
End If

ccIndexNum = ccIndexNum - 2
If Mid(ccNum, ccIndexNum, 1) * 2 > 9 Then
    tempValue = tempValue + (1 + ((Mid(ccNum, ccIndexNum, 1) * 2) Mod 10))
Else
    tempValue = tempValue + (Mid(ccNum, ccIndexNum, 1) * 2)
End If

    'ValidateCCNumber = tempValue

tempValue2 = CInt(Mid(ccNum, 16, 1)) + CInt(Mid(ccNum, 14, 1)) + CInt(Mid(ccNum, 12, 1)) + CInt(Mid(ccNum, 10, 1)) + CInt(Mid(ccNum, 8, 1)) + CInt(Mid(ccNum, 6, 1)) + CInt(Mid(ccNum, 4, 1)) + CInt(Mid(ccNum, 2, 1))

    'ValidateCCNumber = tempValue2

    'ValidateCCNumber = tempValue & "/" & tempValue2

If (CInt(tempValue) + CInt(tempValue2)) Mod 10 = 0 Then
    ValidateCCNumber = 1
Else
    ValidateCCNumber = 0
End If

End Function

Sub fDateYYYYMMDD()

Dim rngR As Range, rngRR As Range
Dim strTemp As String
Dim strTempYear As String, strTempMonth As String, strTempDay As String
Set rngRR = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)

For Each rngR In rngRR
    strTemp = ""
    strTempMonth = Mid(rngR.Value, 5, 2)
    strTempDay = Mid(rngR.Value, 7, 2)
    strTempYear = Mid(rngR.Value, 1, 4)
    strTemp = strTempMonth + "/" + strTempDay + "/" + strTempYear
    
    rngR.Value = strTemp
Next rngR
                
End Sub

Sub FillDownFormula()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Filldown a formula for in column of data.
'   Assumes a data table with headings in the first row,
'   the formula in the second row and is the active cell.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Dim rng As Range
    Dim rngData As Range
    Dim rngFormula As Range
    Dim rowData As Long
    Dim colData As Long

'   Set the ranges
   Set rng = ActiveCell
    Set rngData = rng.CurrentRegion
    
'   Set the row and column variables
   rowData = rngData.CurrentRegion.Rows.Count
    colData = rng.Column

'   Set the formula range and fill down the formula
   Set rngFormula = rngData.Offset(1, colData - 1).Resize(rowData - 1, 1)
    rngFormula.FillDown
End Sub

Function Full_Name(LastName As String, FirstName As String)
    Full_Name = LastName + ", " + FirstName

End Function

Function YesNoMaybe(YourAnswer)
    YesNoMaybe = YourAnswer
End Function



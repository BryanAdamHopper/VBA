VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShowChars 
   Caption         =   "Display Char Codes"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   OleObjectBlob   =   "frmShowChars.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmShowChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    UserForm_Activate
    DoIt
End Sub

Private Sub btnAbout_Click()
    frmAbout.Show
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDeleteChar_Click()
    Dim S As String
    Dim T As String
    Dim n As Long
    
    If ActiveCell.HasFormula = True Then
        Exit Sub
    End If
    
    S = ActiveCell.Text
    n = CLng(Me.cbxStart.Value)
    T = Left(S, n - 1) & Mid(S, n + 1)
    On Error Resume Next
    ActiveCell.Value = T
    DoIt
End Sub

Private Sub cbxStart_Change()
    DoIt
End Sub

Private Sub chkHex_Click()
    DoIt
    Me.cbxStart.SetFocus
End Sub
Private Sub DoIt()
Dim S As String
Dim H As String
Dim l As String
Dim HC As String
Dim SP As String
Dim n As Long
Dim C As String


Me.lblCell.Caption = "Cell: " & ActiveCell.Address(False, False)
Me.lblWorkbook.Caption = "Workbook: " & ActiveWorkbook.FullName
Me.lblWorksheet.Caption = "Worksheet: " & ActiveSheet.Name


If Me.chkHex.Value Then
    Me.lblCode.Caption = "Hex"
Else
    Me.lblCode.Caption = "Dec"
End If

Me.Caption = "Character Codes For Cell: " & ActiveCell.Address(False, False)
Me.lblS.Caption = ""
Me.lblSpec.Caption = ""
Me.lblH.Caption = ""
Me.lblN.Caption = ""

If Len(ActiveCell.Text) > 0 Then
    If Asc(Left(ActiveCell.Text, 1)) <= 32 Then
        If Asc(Right(ActiveCell.Text, 1)) <= 32 Then
            Me.lblLeadTrail.Caption = "Leading and trailing hidden characters or spaces found."
        Else
            Me.lblLeadTrail.Caption = "Leading hidden characters or spaces found."
        End If
    Else
        If Asc(Right(ActiveCell.Text, 1)) <= 32 Then
            Me.lblLeadTrail.Caption = "Trailing hidden characters or spaces found."
        Else
            Me.lblLeadTrail.Caption = ""
        End If
    End If
End If

Me.lblPrefixChar.Caption = vbNullString
If ActiveCell.PrefixCharacter = "'" Then
    Me.lblPrefixChar.Caption = "Cell has an apostrophe prefix character."
    Me.lblPrefixChar.ForeColor = vbRed
End If

Me.lblHasFormula.Caption = vbNullString
If ActiveCell.HasFormula = True Then
    Me.lblHasFormula.Caption = "Cell has a formula. Result is displayed."
    If ActiveCell.HasArray = True Then
        Me.lblHasFormula.Caption = "Cell has a array formula. Result is displayed."
    End If
End If
    
If TypeOf Selection Is Excel.Range Then
    If Selection.Cells.Count > 1 Then
        Me.lblMultiCell.Caption = "Multiple cells selected. Value is cell 1."
    Else
        Me.lblMultiCell.Caption = vbNullString
    End If
End If

    

    
    
    
    
On Error Resume Next
S = " "
SP = " "
If Len(ActiveCell.Text) > 0 Then
    For n = CInt(Me.cbxStart.Value) To Application.Min(CInt(Me.cbxStart.Value) + 23, Len(ActiveCell.Text))
        C = Mid(ActiveCell.Text, n, 1)
        l = l & Format(n, "000 ")
        If (Asc(C) <> 10) And (Asc(C) <> 13) Then
            S = S & C & "   "
        ElseIf Asc(C) = 9 Then
            S = S & Space(4)
        Else
            S = S & Chr(2) & "   "
        End If
        Select Case CInt(Asc(C))
            Case 1 To 31, 160, 129 To 255
                SP = SP & "^   "
'            Case 0: SP = SP & "NUL "
'            Case 1: SP = SP & "SOH "
'            Case 2: SP = SP & "STX "
'            Case 3: SP = SP & "ETX "
'            Case 4: SP = SP & "EOT "
'            Case 5: SP = SP & "ENQ "
'            Case 6: SP = SP & "ACK "
'            Case 7: SP = SP & "BEL "
'            Case 8: SP = SP & "BS  "
'            Case 9: SP = SP & "HT  "
'            Case 10: SP = SP & "LF  "
'            Case 11: SP = SP & "VT  "
'            Case 12: SP = SP & "FF  "
'            Case 13: SP = SP & "CR  "
'            Case 14: SP = SP & "SO  "
'            Case 15: SP = SP & "SI  "
'            Case 16: SP = SP & "SLE "
'            Case 17: SP = SP & "CS1 "
'            Case 18: SP = SP & "DC2 "
'            Case 19: SP = SP & "DC3 "
'            Case 20: SP = SP & "DC4 "
'            Case 21: SP = SP & "NAK "
'            Case 22: SP = SP & "SYN "
'            Case 23: SP = SP & "ETB "
'            Case 24: SP = SP & "CAN "
'            Case 25: SP = SP & "EM  "
'            Case 26: SP = SP & "SLB "
'            Case 27: SP = SP & "ESC "
'            Case 28: SP = SP & "FS  "
'            Case 29: SP = SP & "GS  "
'            Case 30: SP = SP & "RS  "
'            Case 31: SP = SP & "US  "
'            Case 32: SP = SP & "sp "
            Case Else
                SP = SP & "    "
        End Select
        
        If Me.chkHex.Value Then
            HC = "x" & Hex(Asc(C))
            HC = IIf(Len(HC) = 1, "x0" & HC, HC)
            H = H & HC & Application.WorksheetFunction.Choose(Len(HC), "   ", "  ", " ")
        Else
            H = H & Format(Asc(C), "000 ")
        End If
        
    Next n
    Me.lblN.Caption = l
    Me.lblH.Caption = H
    Me.lblS.Caption = S
    Me.lblSpec.Caption = SP
End If
End Sub


Private Sub lblPrefixChar_Click()

End Sub

Private Sub UserForm_Activate()
    Dim n As Long
    With Me.cbxStart
        .Clear
        For n = 1 To Len(ActiveCell.Text)
            .AddItem Format(n, "##0")
        Next n
        If .ListCount > 0 Then
            .ListIndex = 0
        End If
    End With
    
    DoIt
End Sub

Private Sub UserForm_Initialize()
    Set App = Application
    
End Sub


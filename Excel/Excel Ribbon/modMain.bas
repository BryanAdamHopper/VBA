Attribute VB_Name = "modMain"
Option Explicit

Public Const C_TAG = "__HEXVIEW__"
Public Const C_VERISON = "3.2"


Sub ShowTheForm()
#If VBA6 Then
    frmShowChars.Show vbModeless
#Else
    frmShowChars.Show
#End If
End Sub
Sub CreateMenuItem()
DeleteMenuItem
With Application.CommandBars("Worksheet Menu Bar").Controls("View").Controls.Add(Type:=msoControlButton, temporary:=True)
    .Caption = "&View Cell Contents"
    .OnAction = "'" & ThisWorkbook.Name & "'!ShowTheForm"
    .Tag = C_TAG
End With
    
With Application.CommandBars("Cell").Controls.Add(Type:=msoControlButton, temporary:=True)
    .Caption = "&View Cell Content"
    .OnAction = "'" & ThisWorkbook.Name & "'!ShowTheForm"
    .Tag = C_TAG
End With
    
End Sub
Sub DeleteMenuItem()
    On Error Resume Next
    Dim C As Office.CommandBarControl
    Set C = Application.CommandBars.FindControl(Tag:=C_TAG)
    Do Until C Is Nothing
        C.Delete
        Set C = Application.CommandBars.FindControl(Tag:=C_TAG)
    Loop
    For Each C In Application.CommandBars("Cell").Controls
        If StrComp(C.Caption, "&View Cell Content", vbTextCompare) = 0 Then
            C.Delete
        End If
    Next C
End Sub

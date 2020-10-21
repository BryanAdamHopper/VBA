VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "About Cell View"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private Sub btnOK_Click()
Unload Me
End Sub


Private Sub lblAbout_Click()

End Sub

Private Sub lblMail_Click()
    ThisWorkbook.FollowHyperlink "mailto:chip@cpeasron.com?subject=Cell Viewer"
End Sub

Private Sub lblThisURL_Click()
    ThisWorkbook.FollowHyperlink "http://www.cpearson.com/Excel/CellView.aspx"
End Sub

Private Sub lblUrl_Click()
    ThisWorkbook.FollowHyperlink Address:="http://www.cpearson.com/excel"
End Sub


Private Sub UserForm_Initialize()
    Dim S As String
#If VBA7 And Win64 Then
    S = "64-Bit Excel"
#Else
    S = "32-Bit Excel"
#End If
    Me.lblAbout.Caption = "CellView Cell Contents Viewer For Excel" & vbNewLine & _
        "Written in 2002 (major revisions in 2009 and 2012) by Chip Pearson at Pearson Software Consulting, LLC " & vbNewLine & vbNewLine & _
        "© Copyright 2002-2013 Charles H Pearson." & vbNewLine & _
        "chip@cpearson.com     www.cpearson.com" & vbNewLine & vbNewLine & _
        "File Location: " & ThisWorkbook.FullName & vbNewLine & _
        "    File Date: " & Format(FileDateTime(ThisWorkbook.FullName), "dd-MMM-yyyy") & vbNewLine & _
        "    File Size: " & Format(FileLen(ThisWorkbook.FullName) / 1024, "#,##0") & " KB" & vbNewLine & vbNewLine & _
        "Excel Version: " & CStr(Application.Version) & "  Build: " & CStr(Application.Build) & vbNewLine & _
        "   " & S
        
    Me.lblVersion.Caption = "Version " & C_VERISON
    Me.lblUrl.Caption = "www.cpearson.com/Excel"
    Me.lblMail.Caption = "chip@cpearson.com"
    Me.lblThisURL.Caption = "www.cpearson.com/Excel/CellView.aspx"
        
        
End Sub

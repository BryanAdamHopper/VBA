VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRunMacro 
   Caption         =   "Choose Macro"
   ClientHeight    =   570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   OleObjectBlob   =   "frmRunMacro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRunMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnRun_Click()

    Dim MacroAction As String
    MacroAction = cbMacro.Text

    Select Case MacroAction
    
        Case "Remaining Sessions"
            'Call Function RemSessions()
            'frmRunMacro.Hide
            RemSessions
        
        Case "Products"
            'Call Function Products1()
            'frmRunMacro.Hide
            Products1
            
        Case "Undo Last Macro - Alpha"
            'Undo Last Macro
            frmRunMacro.UndoAction
            
        Case Else
            temp1 = MsgBox("Please select a macro to run first!", vbOKOnly)
            
    End Select

End Sub

Private Sub UserForm_Activate()

    With cbMacro
        .AddItem "Products"
        .AddItem "Remaining Sessions"
        .AddItem "Undo Last Macro"
    End With

End Sub


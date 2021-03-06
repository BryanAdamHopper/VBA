Attribute VB_Name = "Module4"
Sub ContactLogs()

Dim MsgBoxText As String
    For x = 1 To Application.Sheets.Count
        MsgBoxText = MsgBoxText & x & " - " & Application.Sheets(x).Name & Chr(13)
    Next x
    mySourceSheet = Application.InputBox(prompt:=MsgBoxText, Type:=1)
    If mySourceSheet = 0 Then
        Exit Sub
    End If

Dim WS_Main As String
    WS_Main = "ControlPanel"
Dim WS_Source As String
    'WS_Source = "Source"
    WS_Source = Application.Sheets(mySourceSheet).Name
Dim WS_Output As String
    'WS_Output = "Output"
    WS_Output = "Output" & Format(DateTime.Now, "hhmmss")
    Worksheets.Add.Name = WS_Output
Dim ColHeaderCount As Double
    ColHeaderCount = Application.WorksheetFunction.CountA(Worksheets(WS_Source).Range("A1:Z1"))
Dim colLetter() As String
    colLetter() = Split("A|B|C|D|E|F|G|H|I|J|K|L|M|N|O|P|Q|R|S|T|U|V|W|X|Y|Z", "|", -1, vbBinaryCompare)
    'A = 0, B = 1, etc...
Dim DataCount As Double
    DataCount = Application.WorksheetFunction.CountA(Worksheets(WS_Source).Range("A:A"))
    
TimeStart = Now

'############################################################################################
'# - Input -
'###########
Dim TblBody(1048576, 25) As String  'Allows for max number of logs. Must be a constant variable
'Dim TblBody(300000, 2) As String   'Limited range for test as size must be a constant
Dim ContactLog As String
    
Worksheets(WS_Source).Columns("A:Z").AutoFit
    
    ContactLogHeader = "<table><tbody><tr>"
    For Z = 1 To (ColHeaderCount - 1)
        ContactLogHeader = ContactLogHeader & "<td>" & Worksheets(WS_Source).Range(colLetter(Z) & 1).Text & "</td>"
    Next Z
    ContactLogHeader = ContactLogHeader & "</tr>"
    ContactLog = ContactLogHeader
    w = 1
    
    For x = 2 To DataCount
        DoEvents
        'Application.StatusBar = "Step 1 of 2 - Injecting Data: " & Format(x / DataCount, "0%")
        Application.StatusBar = "Step 1 of 2 - Injecting Data: " & Format(x / DataCount, "0%") & " | Duration(sec): " & DateDiff("s", TimeStart, Now)
        
        TblBody(w, 1) = Worksheets(WS_Source).Range("A" & x).Text 'repeats more than it needs to
        ContactLog = ContactLog & "<tr>"
        
        For y = 1 To (ColHeaderCount - 1)
            ContactLog = ContactLog _
                    & "<td>" & Worksheets(WS_Source).Range(colLetter(y) & x).Text & "</td>"
        Next y
        ContactLog = ContactLog & "</tr>"
        
        'combine like clientID's or move on to new ones
        If Worksheets(WS_Source).Range("A" & x).Text = Worksheets(WS_Source).Range("A" & x + 1).Text Then
              'w = w (do nothing)
        Else
            ContactLog = ContactLog & "</tbody></table>"
            TblBody(w, 2) = ContactLog
            w = w + 1
            ContactLog = ContactLogHeader
        End If
    Next x

'############################################################################################
'# - Output -
'#
'# Note: Only the first 32,767 characters will be placed into the cell. Need to add a fix to break this up.
'###########

    For x = 1 To w 'w = Total contact logs
        DoEvents
        'Application.StatusBar = "Step 3 of 3 - Output Data: " & Format(x / DataCount, "0%")
        Application.StatusBar = "Step 2 of 2 - Output Data: " & Format(x / w, "0%") & " | Duration(sec): " & DateDiff("s", TimeStart, Now)
    
        'Removed 'For' loop here to limit access to colLetter.
                'Besides, it's only 2 columns...
        Worksheets(WS_Output).Range("A" & x).Value = TblBody(x, 1)
        Worksheets(WS_Output).Range("B" & x).Value = TblBody(x, 2)
    Next x
    
   Application.StatusBar = False
    
End Sub




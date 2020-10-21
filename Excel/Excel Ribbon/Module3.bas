Attribute VB_Name = "Module3"
Sub Abc()
'
' Abc Macro
'

    ActiveCell.FormulaR1C1 = "A"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "B"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "C"
    Range("A1:A3").Select
    Selection.AutoFill Destination:=Range("A1:A94"), Type:=xlFillDefault
    Range("A1:A94").Select
    ActiveWindow.SmallScroll Down:=-135
    Selection.Copy
    Range("B1").Select
    ActiveSheet.Paste
    
End Sub

Sub CIDRSS()
'
' CIDRSS Macro
'

    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "RSSID"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "ClientID"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("A2:B3").Select
    
    Columns("A:B").EntireColumn.AutoFit
    
End Sub

Sub Products()
'
' Products Macro
'

'Start Products code.
    Columns("A:I").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
               
    Range("I1").Select
        ActiveCell.FormulaR1C1 = "Location"
        Range("I2").Select
           ActiveCell.FormulaR1C1 = "1"
        Range("I3").Select
           ActiveCell.FormulaR1C1 = "1"

    Range("H1").Select
        ActiveCell.FormulaR1C1 = "SupplierID"
        Range("H2").Select
            ActiveCell.FormulaR1C1 = "100"
        Range("H3").Select
            ActiveCell.FormulaR1C1 = "100"

    Range("G1").Select
        ActiveCell.FormulaR1C1 = "ProductID"
        Range("G2").Select
            ActiveCell.FormulaR1C1 = "50001"
        Range("G3").Select
            ActiveCell.FormulaR1C1 = "50002"

    Range("F1").Select
        ActiveCell.FormulaR1C1 = "ProductGroupID"
        Range("F2").Select
            ActiveCell.FormulaR1C1 = "50001"
        Range("F3").Select
            ActiveCell.FormulaR1C1 = "50002"

    Range("E1").Select
        ActiveCell.FormulaR1C1 = "DescriptionID"
        Range("E2").Select
            ActiveCell.FormulaR1C1 = "50001"
        Range("E3").Select
            ActiveCell.FormulaR1C1 = "50002"

    Range("D1").Select
        ActiveCell.FormulaR1C1 = "ColorID"
        Range("D2").Select
            ActiveCell.FormulaR1C1 = "1"
        Range("D3").Select
            ActiveCell.FormulaR1C1 = "1"
    
    Range("C1").Select
        ActiveCell.FormulaR1C1 = "SizeID"
        Range("C2").Select
            ActiveCell.FormulaR1C1 = "1"
        Range("C3").Select
            ActiveCell.FormulaR1C1 = "1"

    Range("B1").Select
        ActiveCell.FormulaR1C1 = "CategoryID"
        Range("B2").Select
            ActiveCell.FormulaR1C1 = "99999"
        Range("B3").Select
            ActiveCell.FormulaR1C1 = "99999"

    Range("A1").Select
        ActiveCell.FormulaR1C1 = "SubCategoryID"
    
    
    Columns("A:I").EntireColumn.AutoFit
    Range("B2:I3").Select
    
End Sub

Sub RemSessions()
'
' Remaining Sessions Macro
'
    Columns("A:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Count"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "NumClasses"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "RealRemaining"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "ActiveDate"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Type"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "TyprGroup"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "ProductID"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "ClientID"
    
    Columns("A:H").EntireColumn.AutoFit
    
End Sub

Sub VisitHistory()
'
' Visit History Macro
'
    Columns("A:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove


    Range("M1").Select
    ActiveCell.FormulaR1C1 = "IsPast"
        Range("M2").Select
        ActiveCell.FormulaR1C1 = "IF(NOW()>C2,1,0)"

    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Value"
        Range("L2").Select
        ActiveCell.FormulaR1C1 = "1"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Missed"
        Range("K2").Select
        ActiveCell.FormulaR1C1 = "0"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Cancelled"
        Range("J2").Select
        ActiveCell.FormulaR1C1 = "0"
        
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "TypeGroup"
        Range("I2").Select
        ActiveCell.FormulaR1C1 = "tblvisittypes.Typegroup" 'Typegroup = tblvisittypes.Typegroup
        
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "VisitType"
        Range("H2").Select
        ActiveCell.FormulaR1C1 = "tblvisittypes.typeID" 'VisitType = tblvisittypes.typeID
        
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "TypeTaken"
        Range("G2").Select
        ActiveCell.FormulaR1C1 = "VISITNAME"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "MyEndTime"
        Range("F2").Select
        ActiveCell.FormulaR1C1 = "13:00"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "MyStartTime"
        Range("E2").Select
        ActiveCell.FormulaR1C1 = "12:00"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "ClassTime"
        Range("D2").Select
        ActiveCell.FormulaR1C1 = "12:00"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "ClassDate"
        Range("C2").Select
        ActiveCell.FormulaR1C1 = "10/11/2012"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "TrainerID"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "ClientID"
    
    Columns("A:M").EntireColumn.AutoFit
    
End Sub


Sub States()

Selection.Replace What:="Alabama", Replacement:="AL", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Alaska", Replacement:="AK", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Arizona", Replacement:="AZ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Arkansas", Replacement:="AR", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="California", Replacement:="CA", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Colorado", Replacement:="CO", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Connecticut", Replacement:="CT", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Delaware", Replacement:="DE", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Florida", Replacement:="FL", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Georgia", Replacement:="GA", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Hawaii", Replacement:="HI", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Idaho", Replacement:="ID", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Illinois", Replacement:="IL", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Indiana", Replacement:="IN", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Iowa", Replacement:="IA", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Kansas", Replacement:="KS", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Kentucky", Replacement:="KY", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Louisiana", Replacement:="LA", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Maine", Replacement:="ME", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Maryland", Replacement:="MD", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Massachusetts", Replacement:="MA", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Michigan", Replacement:="MI", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Minnesota", Replacement:="MN", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Mississippi", Replacement:="MS", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Missouri", Replacement:="MO", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Montana", Replacement:="MT", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Nebraska", Replacement:="NE", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Nevada", Replacement:="NV", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="New Hampshire", Replacement:="NH", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="New Jersey", Replacement:="NJ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="New Mexico", Replacement:="NM", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="New York", Replacement:="NY", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="North Carolina", Replacement:="NC", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="North Dakota", Replacement:="ND", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Ohio", Replacement:="OH", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Oklahoma", Replacement:="OK", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Oregon", Replacement:="OR", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Pennsylvania", Replacement:="PA", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Rhode Island", Replacement:="RI", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="South Carolina", Replacement:="SC", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="South Dakota", Replacement:="SD", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Tennessee", Replacement:="TN", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Texas", Replacement:="TX", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Utah", Replacement:="UT", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Vermont", Replacement:="VT", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Virginia", Replacement:="VA", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Washington", Replacement:="WA", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="West Virginia", Replacement:="WV", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Wisconsin", Replacement:="WI", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Wyoming", Replacement:="WY", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="American Samoa", Replacement:="AS", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="District of Columbia", Replacement:="DC", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Federated States of Micronesia", Replacement:="FM", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Guam", Replacement:="GU", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Marshall Islands", Replacement:="MH", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Northern Mariana Islands", Replacement:="MP", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Palau", Replacement:="PW", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Puerto Rico", Replacement:="PR", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Virgin Islands", Replacement:="VI", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Armed Forces Africa", Replacement:="AE", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Armed Forces Americas", Replacement:="AA", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Armed Forces Canada", Replacement:="AE", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Armed Forces Europe", Replacement:="AE", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Armed Forces Middle East", Replacement:="AE", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Armed Forces Pacific", Replacement:="AP", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

End Sub

Sub ReplaceMonthText()

Selection.Replace What:="January", Replacement:="1", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="February", Replacement:="2", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="March", Replacement:="3", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="April", Replacement:="4", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="May", Replacement:="5", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="June", Replacement:="6", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="July", Replacement:="7", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="August", Replacement:="8", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="September", Replacement:="9", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="October", Replacement:="10", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="November", Replacement:="11", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="December", Replacement:="12", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Jan", Replacement:="1", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Feb", Replacement:="2", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Mar", Replacement:="3", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Apr", Replacement:="4", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="May", Replacement:="5", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Jun", Replacement:="6", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Jul", Replacement:="7", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Aug", Replacement:="8", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Sep", Replacement:="9", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Oct", Replacement:="10", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Nov", Replacement:="11", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Selection.Replace What:="Dec", Replacement:="12", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

End Sub

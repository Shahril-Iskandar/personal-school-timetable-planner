Attribute VB_Name = "Module1"
Sub All()

Call FormatList
Call CalendarTable
Call CalendarData
Call CalendarOutline
Call CalendarFormula
Call NameManager
Call Conditioning

End Sub
Sub FormatList()
Attribute FormatList.VB_ProcData.VB_Invoke_Func = " \n14"
'
'
    Range("A1").Select
    Selection.CurrentRegion.Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$J$66"), , xlYes).Name = _
        "Table1"
    Columns("I:J").Select
    Selection.NumberFormat = "[$-en-US,1]h:mm am/pm;@"
    Columns("I:I").Select
    Selection.TextToColumns Destination:=Range("Table1[[#Headers],[Start Time]]") _
        , DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
    Columns("J:J").Select
    Selection.TextToColumns Destination:=Range("Table1[[#Headers],[End Time]]"), _
        DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter _
        :=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, _
        Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Range("Table1[#Headers]").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Range("F1").Select
    ActiveSheet.Paste
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Unique course code"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Group list"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Select Course Code"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Select Group"
    Range("A2").Select
    ActiveCell.Formula2R1C1 = _
        "=IFERROR(INDEX(Table1[Course Code],MATCH(0,INDEX(COUNTIF(R1C1:R[-1]C,Table1[Course Code]),),0)),"""")"
    Range("A2").Select
    Selection.AutoFill Destination:=Range("A2:A50"), Type:=xlFillDefault
    Range("A2:A50").Select
   
    Range("C2").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=OFFSET($A$2,,,COUNTIF($A$2:$A$40,""?*""))"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Cells(2, 3) = Cells(2, 1)
   
    Range("B2").Select
    ActiveCell.Formula2R1C1 = _
        "=FILTER(FILTER(Table1,Table1[Course Code]=R2C3),{0,0,0,0,1,0,0,0,0,0})"
    Range("B3").Select
    
    Range("D2").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=OFFSET($B$2,,,COUNTIF($B$2:$B$40,""?*""))"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Cells(2, 4) = Cells(2, 2)
    
    Range("F2").Select
    ActiveCell.Formula2R1C1 = _
        "=FILTER(Table1,(Table1[Course Code]=RC[-3])*(Table1[Group]=RC[-2]),""NA"")"
    Columns("A:B").Select
    Selection.EntireColumn.Hidden = True
    Columns("C:O").Select
    Columns("C:O").EntireColumn.AutoFit
    Columns("N:O").Select
    Selection.NumberFormat = "[$-en-US,1]h:mm am/pm;@"
    Columns("F:O").Select
    Selection.ColumnWidth = 14.82
    Range("C1").Select
    Columns("G:G").ColumnWidth = 25.73
    
    Range("F14").Select
    ActiveCell.Formula2R1C1 = _
        "Course Adding:"
    
    Range("F15:O27").Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    'Add Course Button
    ActiveSheet.Buttons.Add(94.5, 188, 67, 32).Select
    Selection.OnAction = "AddCourse"
    Selection.Characters.Text = "Add Course"
    With Selection.Characters(Start:=1, Length:=10).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .ColorIndex = 1
    End With
    
    'Add Clear Courses Button
    ActiveSheet.Buttons.Add(95.5, 249, 65, 32).Select
    Selection.OnAction = "ClearCourses"
    Selection.Characters.Text = "Clear Courses"
    With Selection.Characters(Start:=1, Length:=8).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    
    Range("E2").Select

End Sub

Sub AddCourse()

Dim CopyRng As Range, i As Range

    Range("F2:O2").Select
    Selection.Copy
    Range("F2").End(xlDown).Offset(1, 0).Select
    For Each i In Range("F15:F27")
        If IsEmpty(ActiveCell.Value) = False Then
            ActiveCell.Offset(1, 0).Select
        Else
            Selection.PasteSpecial Paste:=xlPasteValues
            Exit For
        End If
    Next
    
    Application.CutCopyMode = False
'    MsgBox ("Course added")

End Sub

Sub CalendarTable()

    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Calendar"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Day"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Subject"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Start Time"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "End Time"
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$D$14"), , xlYes).Name = _
        "Table2"
    
End Sub

Sub CalendarData()

    Range("A2").Select
    ActiveCell.FormulaR1C1 = "='Sheet2'!R[13]C[10]"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "='Sheet2'!R[13]C[5]"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "='Sheet2'!R[13]C[11]"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "='Sheet2'!R[13]C[11]"
    Columns("C:D").Select
    Selection.NumberFormat = "[$-en-US,1]h:mm am/pm;@"
    Columns("A:D").Select
    Selection.EntireColumn.AutoFit
End Sub
Sub CalendarOutline()
Attribute CalendarOutline.VB_ProcData.VB_Invoke_Func = " \n14"

    Range("F2").Select
    ActiveCell.FormulaR1C1 = "8:00"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "8:10"
    Range("F2:F3").Select
    Selection.AutoFill Destination:=Range("F2:F86"), Type:=xlFillDefault
    Range("F2:F86").Select
    ActiveWindow.SmallScroll Down:=-90
    Columns("F:F").Select
    Selection.NumberFormat = "[$-en-US,1]h:mm am/pm;@"
    Columns("F:F").Select
    'Text to column
    Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "MON"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "TUE"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "WED"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "THU"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "FRI"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "SAT"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "SUN"
    Range("N1").Select
    
End Sub

Sub CalendarFormula()
Attribute CalendarFormula.VB_ProcData.VB_Invoke_Func = " \n14"
'
    Range("G2").Select
    ActiveCell.Formula2R1C1 = _
        "=IFERROR(IF(SUMPRODUCT((Table2[Start Time]=RC6)*(R1C=Table2[Day]))=0, """", INDEX(Table2[Subject], SUMPRODUCT((Table2[Day]=R1C)*(Table2[Start Time]=RC6)*(MATCH(ROW(Table2[Day]), ROW(Table2[Day])))))), """")"
    Range("H2").Select
    ActiveCell.Formula2R1C1 = _
        "=IFERROR(IF(SUMPRODUCT((Table2[Start Time]=RC6)*(R1C=Table2[Day]))=0, """", INDEX(Table2[Subject], SUMPRODUCT((Table2[Day]=R1C)*(Table2[Start Time]=RC6)*(MATCH(ROW(Table2[Day]), ROW(Table2[Day])))))), """")"
    Range("I2").Select
    ActiveCell.Formula2R1C1 = _
        "=IFERROR(IF(SUMPRODUCT((Table2[Start Time]=RC6)*(R1C=Table2[Day]))=0, """", INDEX(Table2[Subject], SUMPRODUCT((Table2[Day]=R1C)*(Table2[Start Time]=RC6)*(MATCH(ROW(Table2[Day]), ROW(Table2[Day])))))), """")"
    Range("J2").Select
    ActiveCell.Formula2R1C1 = _
        "=IFERROR(IF(SUMPRODUCT((Table2[Start Time]=RC6)*(R1C=Table2[Day]))=0, """", INDEX(Table2[Subject], SUMPRODUCT((Table2[Day]=R1C)*(Table2[Start Time]=RC6)*(MATCH(ROW(Table2[Day]), ROW(Table2[Day])))))), """")"
    Range("K2").Select
    ActiveCell.Formula2R1C1 = _
        "=IFERROR(IF(SUMPRODUCT((Table2[Start Time]=RC6)*(R1C=Table2[Day]))=0, """", INDEX(Table2[Subject], SUMPRODUCT((Table2[Day]=R1C)*(Table2[Start Time]=RC6)*(MATCH(ROW(Table2[Day]), ROW(Table2[Day])))))), """")"
    Range("L2").Select
    ActiveCell.Formula2R1C1 = _
        "=IFERROR(IF(SUMPRODUCT((Table2[Start Time]=RC6)*(R1C=Table2[Day]))=0, """", INDEX(Table2[Subject], SUMPRODUCT((Table2[Day]=R1C)*(Table2[Start Time]=RC6)*(MATCH(ROW(Table2[Day]), ROW(Table2[Day])))))), """")"
    Range("M2").Select
    ActiveCell.Formula2R1C1 = _
        "=IFERROR(IF(SUMPRODUCT((Table2[Start Time]=RC6)*(R1C=Table2[Day]))=0, """", INDEX(Table2[Subject], SUMPRODUCT((Table2[Day]=R1C)*(Table2[Start Time]=RC6)*(MATCH(ROW(Table2[Day]), ROW(Table2[Day])))))), """")"
    Range("G2:M2").Select
    Selection.AutoFill Destination:=Range("G2:M86"), Type:=xlFillDefault
    Range("G2:M86").Select
    ActiveWindow.SmallScroll Down:=-84
    Range("G2").Select
End Sub
Sub Conditioning()
Attribute Conditioning.VB_ProcData.VB_Invoke_Func = " \n14"
'
    Range("G2:M86").Select
'Color Fill
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=SUMPRODUCT((G$1=Weekday)*($F2>=Start)*($F2<End))"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False ' Conditioning Macro

'NOT Bottom Border
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=SUMPRODUCT((G$1=Weekday)*($F2=Start))"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Borders(xlLeft)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With Selection.FormatConditions(1).Borders(xlRight)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With Selection.FormatConditions(1).Borders(xlTop)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.FormatConditions(1).StopIfTrue = False

'Top Border Only
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=SUMPRODUCT((G$1=Weekday)*($F2=End))"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Borders(xlTop)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.FormatConditions(1).StopIfTrue = False

'Both Sides Border
    ActiveWindow.SmallScroll Down:=-78
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=SUMPRODUCT((G$1=Weekday)*($F2>=Start)*($F2<End))"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Borders(xlLeft)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.FormatConditions(1).Borders(xlRight)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
End Sub
Sub NameManager()
Attribute NameManager.VB_ProcData.VB_Invoke_Func = " \n14"
'
    ActiveWorkbook.Names.Add Name:="Start", RefersToR1C1:= _
        "=OFFSET(Calendar!R2C3,0,0,MATCH(""ZZZZZZZZZZZ"",Calendar!C1)-1)"
    ActiveWorkbook.Names("Start").Comment = ""
    ActiveWorkbook.Names.Add Name:="End", RefersToR1C1:= _
        "=OFFSET(Calendar!R2C4,0,0,MATCH(""ZZZZZZZZZZZ"",Calendar!C1)-1)"
    ActiveWorkbook.Names("End").Comment = ""
    ActiveWorkbook.Names.Add Name:="Weekday", RefersToR1C1:= _
        "=OFFSET(Calendar!R2C1,0,0,MATCH(""ZZZZZZZZZZZ"",Calendar!C1)-1)"
    ActiveWorkbook.Names("Weekday").Comment = ""

End Sub

Sub RemoveCourse()
'Allow user to select which to remove
End Sub

Sub ClearCourses()
'Empty the whole range
    Range("F15:O27").ClearContents
End Sub


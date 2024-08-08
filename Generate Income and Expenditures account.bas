Attribute VB_Name = "Module9"
Sub PL()
    Dim rownum As Long
    Dim Rsum As Long
    Dim Psum As Long
    Dim Depreciation As Long
    Dim K As Integer
    
    
    Depreciation = InputBox("Enter Depreciation ", "Enter Depreciation value of Fixed Assets")
    
    Sheets("P & L").Cells.Clear
    Sheets("P & L").Activate
    Range("C1:E1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Value = "PLANET MARS FOUNDATION"
    
    
    Range("C2:E2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Value = "Registration No. S/66/2016-17"
    
    Range("C3:E3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Value = "INCOME & EXPENDITURES ACCOUNT"
    Range("B4").Value = "RECEIPTS"
    Range("C4").Value = "Rs."
    Range("D4").Value = "PAYMENTS"
    Range("E4").Value = "Rs."
    
    
    Sheets("FinalConsolidation").Activate
    Range("G4").End(xlDown).Select
    rownum = ActiveCell.row
    rownum = rownum - 1
    Sheets("FinalConsolidation").Range("G4:H" & rownum).Copy
    'Activate the destination worksheet
    Sheets("P & L").Activate
    'Select the target range
    Range("B5").Select
    'Paste in the target destination
    ActiveSheet.Paste
    
    Range("B5").End(xlDown).Select
    rownum = ActiveCell.row
    rownum = rownum + 1
    Range("B" & rownum).Value = "Depreciation"
    Range("C" & rownum).Value = Depreciation
    

    Application.CutCopyMode = False
    Range("D5").Value = "Donation received"
    Range("E5").Value = Sheets("Donation").Range("I2")
    
    Sheets("P & L").Activate
    Range("B5").End(xlDown).Select
    rownum = ActiveCell.row
    rownum = rownum + 4
    
    Rsum = Application.WorksheetFunction.sum(Range("C5:C" & rownum - 1))
    Psum = Application.WorksheetFunction.sum(Range("E5:E" & rownum - 1))
    
    If Psum > Rsum Then
        Range("B" & rownum).Value = "Excess Of Income over Expenditure"
        rownum = rownum + 1
        Range("B" & rownum).Formula = "=Sum(C5:C" & rownum - 1 & ")"
        rownum = rownum + 1
        Range("E" & rownum).Formula = "=Sum(E5:E" & rownum - 5 & ")"
        Range("C" & rownum).Formula = Range("E" & rownum).Value
        Range("C" & rownum - 1).Formula = "=(C" & rownum & "- B" & rownum - 1 & ")"
    Else
        Range("E" & rownum).Value = "Excess Of Expenditure over Income"
        rownum = rownum + 1
        Range("C" & rownum).Formula = "=Sum(C5:C" & rownum - 1 & ")"
        rownum = rownum + 1
        Range("E" & rownum).Formula = "=Sum(E5:E" & rownum - 5 & ")"
        Range("C" & rownum).Formula = Range("E" & rownum).Value
        Range("E" & rownum - 1).Formula = "=(C" & rownum - 1 & "- C" & rownum & ")"
    End If
    
    Range("C1:E3").Select
    Selection.Style = "Check Cell"
    Range("B4:E4").Select
    Selection.Style = "Bad"
    Range("B5:B" & rownum - 1).Select
    Selection.Style = "40% - Accent1"
    Range("C5:C" & rownum - 1).Select
    Selection.Style = "Calculation"
    Range("D5:D" & rownum - 1).Select
    Selection.Style = "40% - Accent4"
    Range("E5:E" & rownum - 1).Select
    Selection.Style = "Calculation"
    Range("B" & rownum & ":E" & rownum).Select
    Selection.Style = "Accent4"
    
    
    Range("G4").Value = "Salaries"
    Range("G5").Value = "Opening Accrual (-)"
    Range("G6").Value = "Paid During the year"
    Range("G7").Value = "Closing Accrual"
    Range("G8").Value = "Total Salary Expense"
    
    Range("G10").Value = "Rent"
    Range("G11").Value = "Opening Accrual (-)"
    Range("G12").Value = "Paid During the year"
    Range("G13").Value = "Closing Accrual"
    Range("G14").Value = "Total Rent Expense"
    
    Range("G16").Value = "Utilities"
    Range("G17").Value = "Opening Accrual (-)"
    Range("G18").Value = "Paid During the year"
    Range("G19").Value = "Closing Accrual"
    Range("G20").Value = "Total Utilities Expense"
    
    Sheets("OP&CL").Cells.Clear


    Sheets("April").Range("A2:G3").Copy
    'Activate the destination worksheet
    Sheets("OP&CL").Activate
    'Select the target range
    Range("B2").Select
    'Paste in the target destination
    ActiveSheet.Paste

    Application.CutCopyMode = False
    

    '----------------------------------------------April-----------------------------------------------------------------------
    Sheets("April").Range("A4:G500").Copy
    'Activate the destination worksheet
    Sheets("OP&CL").Activate
    'Select the target range
    Range("B4").Select
    'Paste in the target destination
    ActiveSheet.Paste

    Application.CutCopyMode = False
    
    
    Sheets("OP&CL").Activate
    Range("B2").End(xlDown).Select
    rownum = ActiveCell.row
    rownum = rownum + 1
    
    '----------------------------------------------April-----------------------------------------------------------------------
    Sheets("April").Range("K4:Q500").Copy
    'Activate the destination worksheet
    Sheets("OP&CL").Activate
    'Select the target range
    Range("B" & rownum).Select
    'Paste in the target destination
    ActiveSheet.Paste

    Application.CutCopyMode = False
    
    Sheets("OP&CL").Activate
    Range("B2").End(xlDown).Select
    rownum = ActiveCell.row
    rownum = rownum + 1
    
    '----------------------------------------------April-----------------------------------------------------------------------
    Sheets("April").Range("U4:AA500").Copy
    'Activate the destination worksheet
    Sheets("OP&CL").Activate
    'Select the target range
    Range("B" & rownum).Select
    'Paste in the target destination
    ActiveSheet.Paste

    Application.CutCopyMode = False
    
    
    
    
    
    
    
    
    Sheets("March").Range("A1:G3").Copy
    'Activate the destination worksheet
    Sheets("OP&CL").Activate
    'Select the target range
    Range("K2").Select
    'Paste in the target destination
    ActiveSheet.Paste

    Application.CutCopyMode = False
    Sheets("OP&CL").Range("K2").Value = "."

    '----------------------------------------------March-----------------------------------------------------------------------
    Sheets("March").Range("A3:G500").Copy
    'Activate the destination worksheet
    Sheets("OP&CL").Activate
    'Select the target range
    Range("K4").Select
    'Paste in the target destination
    ActiveSheet.Paste

    Application.CutCopyMode = False
    
    
    Sheets("OP&CL").Activate
    Range("K2").End(xlDown).Select
    rownum = ActiveCell.row
    rownum = rownum + 1
    
    '----------------------------------------------March-----------------------------------------------------------------------
    Sheets("March").Range("K3:Q500").Copy
    'Activate the destination worksheet
    Sheets("OP&CL").Activate
    'Select the target range
    Range("K" & rownum).Select
    'Paste in the target destination
    ActiveSheet.Paste

    Application.CutCopyMode = False
    
    Sheets("OP&CL").Activate
    Range("K2").End(xlDown).Select
    rownum = ActiveCell.row
    rownum = rownum + 1
    
    '----------------------------------------------March-----------------------------------------------------------------------
    Sheets("March").Range("U3:AA500").Copy
    'Activate the destination worksheet
    Sheets("OP&CL").Activate
    'Select the target range
    Range("K" & rownum).Select
    'Paste in the target destination
    ActiveSheet.Paste

    Application.CutCopyMode = False

    Sheets("OP&CL").Activate
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "'OP&CL'!R2C2:R1387C8", Version:=xlPivotTableVersion12).CreatePivotTable _
        TableDestination:="'OP&CL'!R2C19", TableName:="Openpivot", DefaultVersion _
        :=xlPivotTableVersion12
    Sheets("OP&CL").Select
    Cells(2, 19).Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    With ActiveSheet.PivotTables("Openpivot").PivotFields("Details")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Openpivot").PivotFields("Expenses")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("Openpivot").AddDataField ActiveSheet.PivotTables( _
        "Openpivot").PivotFields("Expenses"), "Count of Expenses", xlCount
    With ActiveSheet.PivotTables("Openpivot").PivotFields("Count of Expenses")
        .Caption = "Sum of Expenses"
        .Function = xlSum
    End With
    ActiveWorkbook.ShowPivotTableFieldList = False
    With ActiveSheet.PivotTables("Openpivot").PivotFields("Details")
        .PivotItems("Bank charges").Visible = False
        .PivotItems("Donation").Visible = False
        .PivotItems("Grocery").Visible = False
        .PivotItems("Legal Fees").Visible = False
        .PivotItems("Medical Expenses").Visible = False
        .PivotItems("Opening Balance").Visible = False
        .PivotItems("Postage").Visible = False
        .PivotItems("Ration").Visible = False
        .PivotItems("Stationery").Visible = False
        .PivotItems("Travelling").Visible = False
        .PivotItems("Withdrawal From Bank").Visible = False
        .PivotItems("(blank)").Visible = False
    End With


    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "'OP&CL'!R3C11:R1297C17", Version:=xlPivotTableVersion12).CreatePivotTable _
        TableDestination:="'OP&CL'!R2C22", TableName:="Closepivot", DefaultVersion _
        :=xlPivotTableVersion12
    Sheets("OP&CL").Select
    Cells(2, 22).Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    With ActiveSheet.PivotTables("Closepivot").PivotFields("Details")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Closepivot").PivotFields("Expenses")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("Closepivot").AddDataField ActiveSheet.PivotTables( _
        "Closepivot").PivotFields("Expenses"), "Count of Expenses", xlCount
    With ActiveSheet.PivotTables("Closepivot").PivotFields("Count of Expenses")
        .Caption = "Sum of Expenses"
        .Function = xlSum
    End With
    ActiveWorkbook.ShowPivotTableFieldList = False
    With ActiveSheet.PivotTables("Closepivot").PivotFields("Details")
        .PivotItems("Opening Balance").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    
    Sheets("OP&CL").Activate
    Range("S2").End(xlDown).Select
    rownum = ActiveCell.row
    
    For K = 3 To rownum
        If Range("S" & K).Value = "Rent" Then
            Sheets("P & L").Range("I11").Value = Sheets("OP&CL").Range("T" & K).Value
        ElseIf Range("S" & K).Value = "Salary" Then
            Sheets("P & L").Range("I5").Value = Sheets("OP&CL").Range("T" & K).Value
        ElseIf Range("S" & K).Value = "Utilities" Then
            Sheets("P & L").Range("I17").Value = Sheets("OP&CL").Range("T" & K).Value
        End If
    Next
    
    Sheets("OP&CL").Activate
    Range("V2").End(xlDown).Select
    rownum = ActiveCell.row
    
    For K = 3 To rownum
        If Range("V" & K).Value = "Rent" Then
            Sheets("P & L").Range("I13").Value = Sheets("OP&CL").Range("W" & K).Value
        ElseIf Range("V" & K).Value = "Salary" Then
            Sheets("P & L").Range("I7").Value = Sheets("OP&CL").Range("W" & K).Value
        ElseIf Range("V" & K).Value = "Utilities" Then
            Sheets("P & L").Range("I19").Value = Sheets("OP&CL").Range("W" & K).Value
        End If
    Next
    
    Sheets("R & P").Activate
    Range("E4").End(xlDown).Select
    rownum = ActiveCell.row
    For K = 5 To rownum
        If Range("E" & K).Value = "Rent" Then
            Sheets("P & L").Range("I12").Value = Sheets("R & P").Range("G" & K).Value
        ElseIf Range("E" & K).Value = "Salary" Then
            Sheets("P & L").Range("I6").Value = Sheets("R & P").Range("G" & K).Value
        ElseIf Range("E" & K).Value = "Utilities" Then
            Sheets("P & L").Range("I18").Value = Sheets("R & P").Range("G" & K).Value
        End If
    Next
    
    Sheets("P & L").Activate
    Range("I8").Formula = "=I6+I7-I5"
    Range("I14").Formula = "=I12+I13-I11"
    Range("I20").Formula = "=I18+I19-I17"
    
    
    Sheets("P & L").Activate
    Range("B4").End(xlDown).Select
    rownum = ActiveCell.row
    
     For K = 5 To rownum
        If Range("B" & K).Value = "Rent" Then
            Sheets("P & L").Range("C" & K).Value = Sheets("P & L").Range("I14").Value
        ElseIf Range("B" & K).Value = "Salary" Then
            Sheets("P & L").Range("C" & K).Value = Sheets("P & L").Range("I8").Value
        ElseIf Range("B" & K).Value = "Utilities" Then
            Sheets("P & L").Range("C" & K).Value = Sheets("P & L").Range("I20").Value
        End If
    Next
    
    Range("G4:I4").Select
    Selection.Style = "Good"
    Range("G10:I10").Select
    Selection.Style = "Good"
    Range("G16:I16").Select
    Selection.Style = "Good"
    Range("G5:H8").Select
    Selection.Style = "Bad"
    Range("G11:H14").Select
    Selection.Style = "Bad"
    Range("G17:H20").Select
    Selection.Style = "Bad"
    Range("I5:I7").Select
    Selection.Style = "Calculation"
    Range("I11:I13").Select
    Selection.Style = "Calculation"
    Range("I17:I19").Select
    Selection.Style = "Calculation"
    Range("I8").Select
    Selection.Style = "Check Cell"
    Range("I14").Select
    Selection.Style = "Check Cell"
    Range("I20").Select
    Selection.Style = "Check Cell"
    Range("G8:H8").Select
    Selection.Style = "Neutral"
    Range("G14:H14").Select
    Selection.Style = "Neutral"
    Range("G20:H20").Select
    Selection.Style = "Neutral"
    Columns("H:H").EntireColumn.AutoFit
    Columns("G:G").EntireColumn.AutoFit
End Sub


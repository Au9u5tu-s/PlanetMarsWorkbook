Attribute VB_Name = "Module7"
Sub RP()
    Dim rownum As Long
    Dim finalbal As Long
    Dim bal As Long
    Dim sum As Long
    sum = 0
    
    
    Sheets("R & P").Cells.Clear
    Sheets("R & P").Activate
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
    Selection.Value = "RECEIPTS & PAYMENTS ACCOUNT"
    Range("B4").Value = "RECEIPTS"
    Range("D4").Value = "Rs."
    Range("E4").Value = "PAYMENTS"
    Range("F4").Value = "Rs."
    Range("B5").Value = "To Opening Balance"
    Range("B6").Value = "Cash in Hand"
    Range("B7").Value = "Cash in Corporation bank"
    Range("B8").Value = "Cash in ICICI Bank"
    Range("B10").Value = "To Donation received"
    
    Range("C6").Value = Sheets("April").Range("G3")
    Range("C7").Value = Sheets("April").Range("Q3")
    Range("C8").Value = Sheets("April").Range("AA3")
    
    Range("D8").Formula = "=C6+C7+C8"
    
    Range("D10").Value = Sheets("Donation").Range("I2")
    
    Sheets("FinalConsolidation").Activate
    Range("G4").End(xlDown).Select
    rownum = ActiveCell.row
    rownum = rownum - 1
    Sheets("FinalConsolidation").Range("G4:H" & rownum).Copy
    'Activate the destination worksheet
    Sheets("R & P").Activate
    'Select the target range
    Range("E5").Select
    'Paste in the target destination
    ActiveSheet.Paste

    Application.CutCopyMode = False
    
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Sheets("R & P").Activate
    Range("E5").End(xlDown).Select
    rownum = ActiveCell.row
    rownum = rownum + 2
    
    Range("E" & rownum).Value = "To Closing Balance"
    rownum = rownum + 1
    Range("E" & rownum).Value = "Cash in Hand"
    Sheets("March").Activate
    Range("G2").End(xlDown).Select
    finalbal = ActiveCell.row
    bal = Range("G" & finalbal).Value
    Sheets("R & P").Activate
    Range("F" & rownum).Value = bal
    sum = sum + bal
    rownum = rownum + 1
    
    Range("E" & rownum).Value = "Cash in Corporation bank"
    Sheets("March").Activate
    Range("Q2").End(xlDown).Select
    finalbal = ActiveCell.row
    bal = Range("Q" & finalbal).Value
    Sheets("R & P").Activate
    Range("F" & rownum).Value = bal
    sum = sum + bal
    rownum = rownum + 1
    
    Range("E" & rownum).Value = "Cash in ICICI Bank"
    Sheets("March").Activate
    Range("AA2").End(xlDown).Select
    finalbal = ActiveCell.row
    bal = Range("AA" & finalbal).Value
    Sheets("R & P").Activate
    Range("F" & rownum).Value = bal
    sum = sum + bal
    
    Range("G" & rownum).Value = sum
    rownum = rownum + 1
    Range("G" & rownum).Formula = "=Sum(G5:G" & rownum - 1 & ")"
    Range("D" & rownum).Formula = "=Sum(D8,D10)"
    
    Range("C1:E3").Select
    Selection.Style = "Check Cell"
    Range("B4:G4").Select
    Selection.Style = "Accent2"
    Range("B5:B" & rownum - 1).Select
    Selection.Style = "40% - Accent1"
    Range("C5:C" & rownum - 1).Select
    Selection.Style = "Accent1"
    Range("D8").Select
    Selection.Style = "Calculation"
    Range("D10").Select
    Selection.Style = "Calculation"
    Range("E5:E" & rownum - 1).Select
    Selection.Style = "40% - Accent4"
    Range("F5:F" & rownum - 1).Select
    Selection.Style = "Accent1"
    Range("G5:G" & rownum - 1).Select
    Selection.Style = "Calculation"
    Range("B" & rownum & ":G" & rownum).Select
    Selection.Style = "Accent4"
End Sub

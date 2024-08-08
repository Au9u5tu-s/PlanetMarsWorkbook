Attribute VB_Name = "Module5"
Sub consol()
    Dim rownum As Long

    'Petty Cash Consolidation
    Sheets("Consolidated").Cells.Clear
    Sheets("FinalConsolidation").Cells.Clear
    Sheets("PettyCash").Activate
    
    Range("B4:H3000").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "PettyCash!R2C2:R2288C8", Version:=xlPivotTableVersion12).CreatePivotTable _
        TableDestination:="Consolidated!R3C1", TableName:="PettyPivot", DefaultVersion _
        :=xlPivotTableVersion12
    Sheets("Consolidated").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PettyPivot").PivotFields("Expenses")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PettyPivot").PivotFields("Details")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PettyPivot").AddDataField ActiveSheet.PivotTables( _
        "PettyPivot").PivotFields("Expenses"), "Count of Expenses", xlCount
    With ActiveSheet.PivotTables("PettyPivot").PivotFields("Count of Expenses")
        .Caption = "Sum of Expenses"
        .Function = xlSum
    End With
    Range("G12").Select
    
    
    Sheets("Corporation").Activate
    
    Range("B4:H3000").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Corporation!R2C2:R2288C8", Version:=xlPivotTableVersion12).CreatePivotTable _
        TableDestination:="Consolidated!R3C4", TableName:="CorpPivot", DefaultVersion _
        :=xlPivotTableVersion12
    Sheets("Consolidated").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("CorpPivot").PivotFields("Expenses")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("CorpPivot").PivotFields("Details")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("CorpPivot").AddDataField ActiveSheet.PivotTables( _
        "CorpPivot").PivotFields("Expenses"), "Count of Expenses", xlCount
    With ActiveSheet.PivotTables("CorpPivot").PivotFields("Count of Expenses")
        .Caption = "Sum of Expenses"
        .Function = xlSum
    End With
    Range("G12").Select
    
    
    
     Sheets("ICICI").Activate
    
    Range("B4:H3000").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "ICICI!R2C2:R2288C8", Version:=xlPivotTableVersion12).CreatePivotTable _
        TableDestination:="Consolidated!R3C7", TableName:="IciPivot", DefaultVersion _
        :=xlPivotTableVersion12
    Sheets("Consolidated").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("IciPivot").PivotFields("Expenses")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("IciPivot").PivotFields("Details")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("IciPivot").AddDataField ActiveSheet.PivotTables( _
        "IciPivot").PivotFields("Expenses"), "Count of Expenses", xlCount
    With ActiveSheet.PivotTables("IciPivot").PivotFields("Count of Expenses")
        .Caption = "Sum of Expenses"
        .Function = xlSum
    End With
    Range("G12").Select
    
    Range("J3").Value = "Description"
    Range("K3").Value = "Amount"
    
    Range("A4").End(xlDown).Select
    rownum = ActiveCell.row
    rownum = rownum - 1
    Sheets("Consolidated").Range("A4:B" & rownum).Copy
    'Activate the destination worksheet
    Sheets("Consolidated").Activate
    'Select the target range
    Range("J4").Select
    'Paste in the target destination
    ActiveSheet.Paste

    Application.CutCopyMode = False

    
    Range("D4").End(xlDown).Select
    rownum = ActiveCell.row
    rownum = rownum - 1
    Sheets("Consolidated").Range("D4:E" & rownum).Copy
    'Activate the destination worksheet
    Sheets("Consolidated").Activate
    'Select the target range
    Range("J4").End(xlDown).Offset(1, 0).Select
    'Paste in the target destination
    ActiveSheet.Paste

    Application.CutCopyMode = False
    Range("G4").End(xlDown).Select
    rownum = ActiveCell.row
    rownum = rownum - 1
    Sheets("Consolidated").Range("G4:H" & rownum).Copy
    'Activate the destination worksheet
    Sheets("Consolidated").Activate
    'Select the target range
    Range("J4").End(xlDown).Offset(1, 0).Select
    'Paste in the target destination
    ActiveSheet.Paste

    Application.CutCopyMode = False
    
    Sheets("Consolidated").Activate
    Range("J4").End(xlDown).Select
    rownum = ActiveCell.row
    Range("J3:K" & rownum).Copy
    
    Sheets("FinalConsolidation").Activate
    Range("B3").Select
    'Paste in the target destination
    ActiveSheet.Paste
    Range("B4:C" & rownum).Select
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Range("B3:C3").Select
    Selection.Style = "Neutral"
    
    Sheets("FinalConsolidation").Activate
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "FinalConsolidation!R3C2:R2000C3", Version:=xlPivotTableVersion12). _
        CreatePivotTable TableDestination:="FinalConsolidation!R3C7", TableName:= _
        "FinalPivot", DefaultVersion:=xlPivotTableVersion12
    Sheets("FinalConsolidation").Select
    Cells(3, 7).Select
    With ActiveSheet.PivotTables("FinalPivot").PivotFields("Description")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("FinalPivot").PivotFields("Amount")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("FinalPivot").AddDataField ActiveSheet.PivotTables( _
        "FinalPivot").PivotFields("Amount"), "Count of Amount", xlCount
    With ActiveSheet.PivotTables("FinalPivot").PivotFields("Count of Amount")
        .Caption = "Sum of Amount"
        .Function = xlSum
    End With
    With ActiveSheet.PivotTables("FinalPivot").PivotFields("Description")
        .PivotItems("Withdrawal From Bank").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
        
End Sub

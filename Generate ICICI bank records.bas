Attribute VB_Name = "Module4"
Sub icici()

Sheets("ICICI").Cells.Clear

Sheets("April").Range("U2:AA3").Copy
'Activate the destination worksheet
Sheets("ICICI").Activate
'Select the target range
Range("B2").Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------April-----------------------------------------------------------------------
Sheets("April").Range("U4:AA500").Copy
'Activate the destination worksheet
Sheets("ICICI").Activate
'Select the target range
Range("B4").Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------May-----------------------------------------------------------------------
Sheets("May").Range("U4:AA500").Copy
'Activate the destination worksheet
Sheets("ICICI").Activate
'Select the target range
Range("B3").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------June-----------------------------------------------------------------------
Sheets("June").Range("U4:AA500").Copy
'Activate the destination worksheet
Sheets("ICICI").Activate
'Select the target range
Range("B3").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------July-----------------------------------------------------------------------
Sheets("July").Range("U4:AA500").Copy
'Activate the destination worksheet
Sheets("ICICI").Activate
'Select the target range
Range("B3").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------August-----------------------------------------------------------------------
Sheets("August").Range("U4:AA500").Copy
'Activate the destination worksheet
Sheets("ICICI").Activate
'Select the target range
Range("B3").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------September-----------------------------------------------------------------------
Sheets("September").Range("U4:AA500").Copy
'Activate the destination worksheet
Sheets("ICICI").Activate
'Select the target range
Range("B3").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------October-----------------------------------------------------------------------
Sheets("October").Range("U4:AA500").Copy
'Activate the destination worksheet
Sheets("ICICI").Activate
'Select the target range
Range("B3").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------November-----------------------------------------------------------------------
Sheets("November").Range("U4:AA500").Copy
'Activate the destination worksheet
Sheets("ICICI").Activate
'Select the target range
Range("B3").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------December-----------------------------------------------------------------------
Sheets("December").Range("U4:AA500").Copy
'Activate the destination worksheet
Sheets("ICICI").Activate
'Select the target range
Range("B3").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------January-----------------------------------------------------------------------
Sheets("January").Range("U4:AA500").Copy
'Activate the destination worksheet
Sheets("ICICI").Activate
'Select the target range
Range("B3").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------February-----------------------------------------------------------------------
Sheets("February").Range("U4:AA500").Copy
'Activate the destination worksheet
Sheets("ICICI").Activate
'Select the target range
Range("B3").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------March-----------------------------------------------------------------------
Sheets("March").Range("U4:AA500").Copy
'Activate the destination worksheet
Sheets("ICICI").Activate
'Select the target range
Range("B3").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'Sheets("April").Range("K4:Q500").Copy
'Activate the destination worksheet
'Sheets("PettyCash").Activate
'Select the target range
'Range("B3").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
'ActiveSheet.Paste

Application.CutCopyMode = False

Range("H4").Formula = "=H3+G4-F4"
End Sub





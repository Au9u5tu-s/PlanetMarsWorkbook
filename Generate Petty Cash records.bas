Attribute VB_Name = "Module1"
Sub sbCopyRangeToAnotherSheet()

Sheets("PettyCash").Cells.Clear


Sheets("April").Range("A2:G3").Copy
'Activate the destination worksheet
Sheets("PettyCash").Activate
'Select the target range
Range("B2").Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------April-----------------------------------------------------------------------
Sheets("April").Range("A4:G500").Copy
'Activate the destination worksheet
Sheets("PettyCash").Activate
'Select the target range
Range("B4").Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------May-----------------------------------------------------------------------
Sheets("May").Range("A4:G500").Copy
'Activate the destination worksheet
Sheets("PettyCash").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------June-----------------------------------------------------------------------
Sheets("June").Range("A4:G500").Copy
'Activate the destination worksheet
Sheets("PettyCash").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------July-----------------------------------------------------------------------
Sheets("July").Range("A4:G500").Copy
'Activate the destination worksheet
Sheets("PettyCash").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------August-----------------------------------------------------------------------
Sheets("August").Range("A4:G500").Copy
'Activate the destination worksheet
Sheets("PettyCash").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------September-----------------------------------------------------------------------
Sheets("September").Range("A4:G500").Copy
'Activate the destination worksheet
Sheets("PettyCash").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------October-----------------------------------------------------------------------
Sheets("October").Range("A4:G500").Copy
'Activate the destination worksheet
Sheets("PettyCash").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------November-----------------------------------------------------------------------
Sheets("November").Range("A4:G500").Copy
'Activate the destination worksheet
Sheets("PettyCash").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------December-----------------------------------------------------------------------
Sheets("December").Range("A4:G500").Copy
'Activate the destination worksheet
Sheets("PettyCash").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------January-----------------------------------------------------------------------
Sheets("January").Range("A4:G500").Copy
'Activate the destination worksheet
Sheets("PettyCash").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------February-----------------------------------------------------------------------
Sheets("February").Range("A4:G500").Copy
'Activate the destination worksheet
Sheets("PettyCash").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------March-----------------------------------------------------------------------
Sheets("March").Range("A4:G500").Copy
'Activate the destination worksheet
Sheets("PettyCash").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
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



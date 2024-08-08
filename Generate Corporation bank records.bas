Attribute VB_Name = "Module3"
Sub Corp()

Sheets("Corporation").Cells.Clear

Sheets("April").Range("K2:Q3").Copy
'Activate the destination worksheet
Sheets("Corporation").Activate
'Select the target range
Range("B2").Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------April-----------------------------------------------------------------------
Sheets("April").Range("K4:Q500").Copy
'Activate the destination worksheet
Sheets("Corporation").Activate
'Select the target range
Range("B4").Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------May-----------------------------------------------------------------------
Sheets("May").Range("K4:Q500").Copy
'Activate the destination worksheet
Sheets("Corporation").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------June-----------------------------------------------------------------------
Sheets("June").Range("K4:Q500").Copy
'Activate the destination worksheet
Sheets("Corporation").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------July-----------------------------------------------------------------------
Sheets("July").Range("K4:Q500").Copy
'Activate the destination worksheet
Sheets("Corporation").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------August-----------------------------------------------------------------------
Sheets("August").Range("K4:Q500").Copy
'Activate the destination worksheet
Sheets("Corporation").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False

'----------------------------------------------September-----------------------------------------------------------------------
Sheets("September").Range("K4:Q500").Copy
'Activate the destination worksheet
Sheets("Corporation").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False
'----------------------------------------------October-----------------------------------------------------------------------
Sheets("October").Range("K4:Q500").Copy
'Activate the destination worksheet
Sheets("Corporation").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False
'----------------------------------------------November-----------------------------------------------------------------------
Sheets("November").Range("K4:Q500").Copy
'Activate the destination worksheet
Sheets("Corporation").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False
'----------------------------------------------December-----------------------------------------------------------------------
Sheets("December").Range("K4:Q500").Copy
'Activate the destination worksheet
Sheets("Corporation").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False
'----------------------------------------------January-----------------------------------------------------------------------
Sheets("January").Range("K4:Q500").Copy
'Activate the destination worksheet
Sheets("Corporation").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False
'----------------------------------------------February-----------------------------------------------------------------------
Sheets("February").Range("K4:Q500").Copy
'Activate the destination worksheet
Sheets("Corporation").Activate
'Select the target range
Range("B4").End(xlDown).Offset(1, 0).Select
'Paste in the target destination
ActiveSheet.Paste

Application.CutCopyMode = False
'----------------------------------------------March-----------------------------------------------------------------------
Sheets("March").Range("K4:Q500").Copy
'Activate the destination worksheet
Sheets("Corporation").Activate
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




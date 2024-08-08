Attribute VB_Name = "Module2"
Sub donation()
    Sheets("Donation").Cells.Clear

    Dim K As Long
    Dim J As Long
    Dim Lastrow As Long


    Sheets("April").Range("K2:P2").Copy
    'Activate the destination worksheet
    Sheets("Donation").Activate
    'Select the target range
    Range("B2").Select
    'Paste in the target destination
    ActiveSheet.Paste

    Application.CutCopyMode = False
    
    Application.ScreenUpdating = False
    J = 3
    Lastrow = 1000
    
'----------------------------------------------April-----------------------------------------------------------------------
    
    For K = 4 To Lastrow
        If CStr(Sheets("April").Range("M" & K).Value) = "Donation" Then
            Sheets("April").Range("K" & K & ":P" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
            J = J + 1
        End If
        If CStr(Sheets("April").Range("W" & K).Value) = "Donation" Then
            Sheets("April").Range("U" & K & ":Z" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
    Next
'----------------------------------------------may-----------------------------------------------------------------------

    
    For K = 4 To Lastrow
        If CStr(Sheets("May").Range("M" & K).Value) = "Donation" Then
            Sheets("May").Range("K" & K & ":P" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
        If CStr(Sheets("May").Range("W" & K).Value) = "Donation" Then
            Sheets("May").Range("U" & K & ":Z" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
    Next
    
'----------------------------------------------june-----------------------------------------------------------------------
    
    For K = 4 To Lastrow
        If CStr(Sheets("June").Range("M" & K).Value) = "Donation" Then
            Sheets("June").Range("K" & K & ":P" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
        If CStr(Sheets("June").Range("W" & K).Value) = "Donation" Then
            Sheets("June").Range("U" & K & ":Z" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
    Next

    
    
'----------------------------------------------july-----------------------------------------------------------------------
    
    For K = 4 To Lastrow
        If CStr(Sheets("July").Range("M" & K).Value) = "Donation" Then
            Sheets("July").Range("K" & K & ":P" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
        If CStr(Sheets("July").Range("W" & K).Value) = "Donation" Then
            Sheets("July").Range("U" & K & ":Z" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
    Next

'----------------------------------------------August-----------------------------------------------------------------------
    
    For K = 4 To Lastrow
        If CStr(Sheets("August").Range("M" & K).Value) = "Donation" Then
            Sheets("August").Range("K" & K & ":P" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
        If CStr(Sheets("August").Range("W" & K).Value) = "Donation" Then
            Sheets("August").Range("U" & K & ":Z" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
    Next
    

'----------------------------------------------September-----------------------------------------------------------------------
    
    For K = 4 To Lastrow
        If CStr(Sheets("September").Range("M" & K).Value) = "Donation" Then
            Sheets("September").Range("K" & K & ":P" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
        If CStr(Sheets("September").Range("W" & K).Value) = "Donation" Then
            Sheets("September").Range("U" & K & ":Z" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
    Next

'----------------------------------------------October-----------------------------------------------------------------------
    
    For K = 4 To Lastrow
        If CStr(Sheets("October").Range("M" & K).Value) = "Donation" Then
            Sheets("October").Range("K" & K & ":P" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
        If CStr(Sheets("October").Range("W" & K).Value) = "Donation" Then
            Sheets("October").Range("U" & K & ":Z" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
    Next

'----------------------------------------------November-----------------------------------------------------------------------
    
    For K = 4 To Lastrow
        If CStr(Sheets("November").Range("M" & K).Value) = "Donation" Then
            Sheets("November").Range("K" & K & ":P" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
        If CStr(Sheets("November").Range("W" & K).Value) = "Donation" Then
            Sheets("November").Range("U" & K & ":Z" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
    Next

'----------------------------------------------December-----------------------------------------------------------------------
    
    For K = 4 To Lastrow
        If CStr(Sheets("December").Range("M" & K).Value) = "Donation" Then
            Sheets("December").Range("K" & K & ":P" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
        If CStr(Sheets("December").Range("W" & K).Value) = "Donation" Then
            Sheets("December").Range("U" & K & ":Z" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
    Next

'----------------------------------------------January-----------------------------------------------------------------------
    
    For K = 4 To Lastrow
        If CStr(Sheets("January").Range("M" & K).Value) = "Donation" Then
            Sheets("January").Range("K" & K & ":P" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
        If CStr(Sheets("January").Range("W" & K).Value) = "Donation" Then
            Sheets("January").Range("U" & K & ":Z" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
    Next

'----------------------------------------------February-----------------------------------------------------------------------
    
    For K = 4 To Lastrow
        If CStr(Sheets("February").Range("M" & K).Value) = "Donation" Then
            Sheets("February").Range("K" & K & ":P" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
        If CStr(Sheets("February").Range("W" & K).Value) = "Donation" Then
            Sheets("February").Range("U" & K & ":Z" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
    Next

'----------------------------------------------March-----------------------------------------------------------------------
    
    For K = 4 To Lastrow
        If CStr(Sheets("March").Range("M" & K).Value) = "Donation" Then
            Sheets("March").Range("K" & K & ":P" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
        If CStr(Sheets("March").Range("W" & K).Value) = "Donation" Then
            Sheets("March").Range("U" & K & ":Z" & K).Copy
            Sheets("Donation").Activate
            Range("B" & J).Select
            ActiveSheet.Paste

            Application.CutCopyMode = False
            J = J + 1
        End If
    Next


    Application.ScreenUpdating = True
    Sheets("Donation").Activate
    Columns(6).EntireColumn.Delete
    
    Range("I2").Formula = "=Sum(F:F)"
End Sub



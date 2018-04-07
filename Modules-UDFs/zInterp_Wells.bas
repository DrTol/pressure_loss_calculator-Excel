Attribute VB_Name = "zInterp_Wells"
Function Linterp(ByVal KnownYs As Range, ByVal KnownXs As Range, NewX As Variant) As Variant
'******************************************************************************
'***DEVELOPER: Ryan Wells (wellsr.com) *
'***DATE: 03/2016 *
'***DESCRIPTION: 2D Linear Interpolation function that automatically picks *
'*** which range to interpolate between based on the closest *
'*** KnownX value to the NewX value you want to interpolate for. *
'***INPUT: KnownYs - 1D range containing your known Y values. *
'*** KnownXs - 1D range containing your known X values. *
'*** NewX - Cell or number with the X value you want to *
'*** interpolate for. *
'***OUTPUT: The output will be the linear interpolated Y value *
'*** corresponding to the NewX value the user selects. *
'***NOTES: i. KnownYs do not have to be sorted. If the values are *
'*** unsorted, the function will linearly interpolate between the *
'*** two closest values to your NewX (one above, one below). *
'*** ii. KnownXs and KnownYs must be the same dimensions. It is a *
'*** good practice to have the Xs and corresponding Ys beside *
'*** each other in Excel before using Linterp. *
'***FORMULA: Linterp=Y0 + (Y1-Y0)*(NewX-X0)/(X1-X0) *
'***EXAMPLE: =Linterp(A2:A4,B2:B4,C2) *
'******************************************************************************
 
'------------------------------------------------------------------------------
'0. Declare Variables and Initialize Variables
'------------------------------------------------------------------------------
Dim bYRows As Boolean   'Y values are selected in a row (Nx1)
Dim bXRows As Boolean   'X values are selected in a row (Nx1)
Dim DeltaHi As Double   'delta between NewX and KnownXs if Known > NewX
Dim DeltaLo As Double   'delta between NewX and KnownXs if Known < NewX
Dim iHi As Long         'Index position of the closest value above NewX
Dim iLo As Long         'Index position of the closest value below NewX
Dim i As Long           'dummy counter
Dim Y0 As Double, Y1 As Double 'Linear Interpolation Y variables
Dim X0 As Double, X1 As Double 'Linear Interpolation Y variables
iHi = 2147483647
iLo = -2147483648#
DeltaHi = 1.79769313486231E+308
DeltaLo = -1.79769313486231E+308
 
'------------------------------------------------------------------------------
'I. Preliminary Error Checking
'------------------------------------------------------------------------------
'Error 0 - catch all error
On Error GoTo InterpError:
'Error 1 - NewX more than 1 cell selected
If IsArray(NewX) = True Then
    If NewX.count <> 1 Then
        Linterp = "Too many cells in variable NewX."
        Exit Function
    End If
End If
 
'Error 2 - NewX is not a number
If IsNumeric(NewX) = False Then
    Linterp = "NewX is non-numeric."
    Exit Function
End If
 
'Error 3 - dimensions aren't even
If KnownYs.count <> KnownXs.count Or _
   KnownYs.Rows.count <> KnownXs.Rows.count Or _
   KnownYs.Columns.count <> KnownXs.Columns.count Then
    Linterp = "Known ranges are different dimensions."
    Exit Function
End If
 
'Error 4 - known Ys are not Nx1 or 1xN dimensions
If KnownYs.Rows.count <> 1 And KnownYs.Columns.count <> 1 Then
    Linterp = "Known Y's should be in a single column or a single row."
    Exit Function
End If
 
'Error 5 - known Xs are not Nx1 or 1xN dimensions
If KnownXs.Rows.count <> 1 And KnownXs.Columns.count <> 1 Then
    Linterp = "Known X's should be in a single column or a single row."
    Exit Function
End If
 
'Error 6 - Too few known Y cells
If KnownYs.Rows.count <= 1 And KnownYs.Columns.count <= 1 Then
    Linterp = "Known Y's range must be larger than 1 cell"
    Exit Function
End If
 
'Error 7 - Too few known X cells
If KnownXs.Rows.count <= 1 And KnownXs.Columns.count <= 1 Then
    Linterp = "Known X's range must be larger than 1 cell"
    Exit Function
End If
 
'Error 8 - Check for non-numeric KnownYs
If KnownYs.Rows.count > 1 Then
    bYRows = True
    For i = 1 To KnownYs.Rows.count
        If IsNumeric(KnownYs.Cells(i, 1)) = False Then
            Linterp = "One or all Known Y's are non-numeric."
            Exit Function
        End If
    Next i
ElseIf KnownYs.Columns.count > 1 Then
    bYRows = False
    For i = 1 To KnownYs.Columns.count
        If IsNumeric(KnownYs.Cells(1, i)) = False Then
            Linterp = "One or all KnownYs are non-numeric."
            Exit Function
        End If
    Next i
End If
 
'Error 9 - Check for non-numeric KnownXs
If KnownXs.Rows.count > 1 Then
    bXRows = True
    For i = 1 To KnownXs.Rows.count
        If IsNumeric(KnownXs.Cells(i, 1)) = False Then
            Linterp = "One or all Known X's are non-numeric."
            Exit Function
        End If
    Next i
ElseIf KnownXs.Columns.count > 1 Then
    bXRows = False
    For i = 1 To KnownXs.Columns.count
        If IsNumeric(KnownXs.Cells(1, i)) = False Then
            Linterp = "One or all Known X's are non-numeric."
            Exit Function
        End If
    Next i
End If
 
'------------------------------------------------------------------------------
'II. Check for nearest values from list of Known X's
'------------------------------------------------------------------------------
If bXRows = True Then 'check by rows
    For i = 1 To KnownXs.Rows.count 'loop through known Xs
        If KnownXs.Cells(i, 1) <> "" Then
            If KnownXs.Cells(i, 1) > NewX And KnownXs.Cells(i, 1) - NewX < DeltaHi Then 'determine DeltaHi
                DeltaHi = KnownXs.Cells(i, 1) - NewX
                iHi = i
            ElseIf KnownXs.Cells(i, 1) < NewX And KnownXs.Cells(i, 1) - NewX > DeltaLo Then 'determine DeltaLo
                DeltaLo = KnownXs.Cells(i, 1) - NewX
                iLo = i
            ElseIf KnownXs.Cells(i, 1) = NewX Then 'match. just report corresponding Y
                Linterp = KnownYs.Cells(i, 1)
                Exit Function
            End If
        End If
    Next i
Else ' check by columns
    For i = 1 To KnownXs.Columns.count 'loop through known Xs
        If KnownXs.Cells(1, i) <> "" Then
            If KnownXs.Cells(1, i) > NewX And KnownXs.Cells(1, i) - NewX < DeltaHi Then 'determine DeltaHi
                DeltaHi = KnownXs.Cells(1, i) - NewX
                iHi = i
            ElseIf KnownXs.Cells(1, i) < NewX And KnownXs.Cells(1, i) - NewX > DeltaLo Then 'determine DeltaLo
                DeltaLo = KnownXs.Cells(1, i) - NewX
                iLo = i
            ElseIf KnownXs.Cells(1, i) = NewX Then 'match. just report corresponding Y
                Linterp = KnownYs.Cells(1, i)
                Exit Function
            End If
        End If
    Next i
End If
 
'------------------------------------------------------------------------------
'III. Linear interpolate based on the closest cells in the range. Includes minor error handling
'------------------------------------------------------------------------------
If iHi = 2147483647 Or iLo = -2147483648# Then
    Linterp = "NewX is out of range. Cannot linearly interpolate with the given Knowns."
    Exit Function
End If
If bXRows = True Then
    Y0 = KnownYs.Cells(iLo, 1)
    Y1 = KnownYs.Cells(iHi, 1)
    X0 = KnownXs.Cells(iLo, 1)
    X1 = KnownXs.Cells(iHi, 1)
Else
    Y0 = KnownYs.Cells(1, iLo)
    Y1 = KnownYs.Cells(1, iHi)
    X0 = KnownXs.Cells(1, iLo)
    X1 = KnownXs.Cells(1, iHi)
End If
Linterp = Y0 + (Y1 - Y0) * (NewX - X0) / (X1 - X0)
Exit Function
 
'------------------------------------------------------------------------------
'IV. Final Error Handling
'------------------------------------------------------------------------------
InterpError:
    Linterp = "Error Encountered: " & Err.Number & ", " & Err.Description
End Function


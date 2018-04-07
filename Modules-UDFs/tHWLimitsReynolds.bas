Attribute VB_Name = "tHWLimitsReynolds"
Function tReynoldsLimits(rRou_or_C As Double, InputType As String) As Variant
' Reynolds limitations for Hazen-Williams formula
'   based on data given by Diskin,M.H. - The limits of applicability of the Hazen-Williams formula
'   The data by Diskin can be found in the Excel sheet "zDiskinData"
'  by Tol,Hakan Ibrahim from the PhD study at Technical University of Denmark
'   PhD Topic: District Heating in Areas with Low-Energy Houses

' DESRIPTION
'   INPUTS
'    rRou_C     : Relative roughness (eps/D) or roughness coefficient (C), used respectively in Darcy-Weisbach (DW) or Hazen-Williams (HW) formulations
'                 Depends on the user choice of the conversion direction
'    InputType  : as for a given input either 'rRou' as relative rougness or 'C' as Hazen-Williams roughness coefficient

'   OUTPUT as ARRAY - tReynoldsLimits keeping, in order, the following returns;
'    maxRe      : Limitations of HW method as maximum Reynolds number on the given roughness
'    minRe      : Limitations of HW method as minimum Reynolds number on the given roughness

Dim tempResult(1 To 2) As Double

' Calculations as to the input type, either "rRou" or "C"

Select Case InputType
    Case "rRou" ' as a function of relative roughness rRou
        tempResult(1) = Linterp(Range("maxRe_Data"), Range("rRou_Data"), rRou_or_C) 'limit for Maximum Reynolds
        tempResult(2) = Linterp(Range("minRe_Data"), Range("rRou_Data"), rRou_or_C) 'limit for Minimum Reynolds
    Case "C" ' as a function of C
        tempResult(1) = Linterp(Range("maxRe_Data"), Range("Cmod_Data"), rRou_or_C) 'limit for Maximum Reynolds
        tempResult(2) = Linterp(Range("minRe_Data"), Range("Cmod_Data"), rRou_or_C) 'limit for Minimum Reynolds
    Case Else
        MsgBox "You must choose either -rRou- or -C- to indicate the input type"
End Select

tReynoldsLimits = tempResult

End Function



Attribute VB_Name = "DWf_ZigrangSylvester"
Function f_ZigrangSylvester(D As Double, Re As Double, aRou As Double)
' Calculates the Darcy-Weisbach friction factor for pressure loss calculations
'   via explicit equation by Zigrang,D.J & Sylvester, N.D.
'   by Tol,Hakan Ibrahim from the PhD study at Technical University of Denmark
'   PhD Topic: District Heating in Areas with Low-Energy Houses

' INPUTS
'    aRou     : Absolute roughness of pipe          [mm]
'    D        : Inner diameter of the pipe          [mm]
'    Re       : Reynolds Number                     [-]

' Checking algorithm limitations
If Re < 4000 Or Re > 100000000 Then
    'MsgBox "Error: Zigrang&Sylvester algorithm is valid for a Reynold range as in 4000<Re<1e8"
    f_ZigrangSylvester = CVErr(xlErrNA)
End If
If aRou / D < 0.00004 Or aRou / D > 0.05 Then
    'MsgBox "Error: Zigrang&Sylvester algorithm is valid for a relative roughness range as in 4e-5<eps/D<0.05"
    f_ZigrangSylvester = CVErr(xlErrNA)
End If

' Fasten your seat belts - Formulation in Run
f_ZigrangSylvester = (-2 * LogBase(((aRou / D) / 3.7 - (5.02 / Re) * LogBase(((aRou / D) - 5.02 / Re * LogBase((((aRou / D) / 3.7) + 13 / Re), 10)), 10)), 10)) ^ -2

End Function

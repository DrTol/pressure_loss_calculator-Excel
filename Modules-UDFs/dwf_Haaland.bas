Attribute VB_Name = "DWf_Haaland"
Function f_Haaland(D As Double, Re As Double, aRou As Double) As Double
' Calculates the Darcy-Weisbach friction factor for pressure loss calculations
'   via explicit equation by Haaland,S.E.
'   by Tol,Hakan Ibrahim from the PhD study at Technical University of Denmark
'   PhD Topic: District Heating in Areas with Low-Energy Houses

' INPUTS
'    aRou     : Absolute roughness of pipe          [mm]
'    D        : Inner diameter of the pipe          [mm]
'    Re       : Reynolds Number                     [-]

' Checking algorithm limitations
If Re < 4000 Or Re > 100000000 Then
    'MsgBox "Error: Haaland algorithm is valid for a Reynold range as in 4000<Re<1e8"
    f_Haaland = CVErr(xlErrNA)
End If
If aRou / D < 0.000001 Or aRou / D > 0.05 Then
    'MsgBox "Error: Haaland algorithm is valid for a relative roughness range as in 1e-6<eps/D<0.05')"
    f_Haaland = CVErr(xlErrNA)
End If

' Fasten your seat belts - Formulation in Run
f_Haaland = (-1.8 * LogBase((((aRou / D) / 3.7) ^ 1.11 + 6.9 / Re), 10)) ^ -2

End Function

Attribute VB_Name = "DWf_Moody"
Function f_Moody(D As Double, Re As Double, aRou As Double) As Double
' Calculates the Darcy-Weisbach friction factor for pressure loss calculations
'   based on Moody correlation
'   by Tol,Hakan Ibrahim from the PhD study at Technical University of Denmark
'   PhD Topic: District Heating in Areas with Low-Energy Houses

' INPUTS
'    aRou     : Absolute roughness of pipe          [mm]
'    D        : Inner diameter of the pipe          [mm]
'    Re       : Reynolds Number                     [-]

' Checking algorithm limitations
If Re < 4000 Or Re > 500000000 Then
    'MsgBox "Error: Moody algorithm is valid for a Reynold range as in 4000<Re<5e8"
    f_Moody = CVErr(xlErrNA)
End If
If aRou / D > 0.01 Then
    'MsgBox "Error: Moody algorithm is valid for a relative roughness range as eps/D<0.01"
    f_Moody = CVErr(xlErrNA)
End If

' Fasten your seat belts - Formulation in Run
f_Moody = 0.0055 * (1 + (20000 * (aRou / D) + 1000000 / Re) ^ (1 / 3))

End Function

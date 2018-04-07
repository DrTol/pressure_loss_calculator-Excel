Attribute VB_Name = "DWf_ColebrookWhite"
Function f_ColebrookWhite(D As Double, Re As Double, aRou As Double, Optional ByVal fTol As Double = 0.01, Optional ByVal MaxIter As Double = 1000) As Double
' Calculates the Darcy-Weisbach friction factor for pressure loss calculations
'   via solving the implicit Colebrook&White expression, details in https://doi.org/10.1098/rspa.1937.0150
'   by Tol,Hakan Ibrahim from the PhD study at Technical University of Denmark
'  PhD Topic: District Heating in Areas with Low-Energy Houses

' DESCRIPTION
'   INPUTS
'    aRou     : Absolute roughness of pipe          [mm]
'    D        : Inner diameter of the pipe          [mm]
'    Re       : Reynolds Number                     [-]

'   INPUTS (for Iteration)
'    fTol     : Termination Tolerance(Iteration)    [-]
'    MaxIter  : Max. limit (Iteration)              [-]

' Initializing the Iteration
Err = 10    ' Iteration error
IterNum = 0 ' Iteration steps number

'  Initial estimate (making use of SwameeJain algorithm)
X0 = Rnd 'Random number
'X0 = f_Clamond(Re, aRou / D)

' Fasten your seat belts, iteration starts

Do While (Err > fTol And IterNum < MaxIter)
    IterNum = IterNum + 1
    X1 = (2 * LogBase(((aRou / D) / 3.7 + 2.51 / (Re * X0 ^ 0.5)), 10)) ^ (-2)
    Err = Abs(X1 - X0)
    X0 = X1
Loop

f_ColebrookWhite = X1

End Function

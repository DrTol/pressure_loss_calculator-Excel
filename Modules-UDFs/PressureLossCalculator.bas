Attribute VB_Name = "PressureLossCalculator"
Function PressureLoss(L As Double, D As Double, aRou As Double, mFlow As Double, T As Double, P As Double, Optional ByVal Solver As String = "Darcy-Weisbach", Optional ByVal Algorithm As String = "Clamond", Optional ByVal fTol As Double = 0.01, Optional ByVal MaxIter As Double = 1000)

' Pipe pressure loss calculator for a circular pipe, full flow water (SI Units)
'   by Tol,Hakan Ibrahim from the PhD study at Technical University of Denmark
'   PhD Topic: District Heating in Areas with Low-Energy Houses

'   INPUTS
'    L          : Length of the pipe segment    [m]
'    D          : Inner Diameter of the pipe    [mm]
'    aRou       : Absolute Roughness of pipe    [mm]
'    mFlow      : Mass Flow                     [kg/s]
'    T          : Water Temperature             [ºC]
'    P          : Hydraulic Static Pressure     [bar]
'   OPTIONAL INPUTS
'    Solver     : as frictional pipe pressure loss formulas, select from:
'                 1-   "Darcy-Weisbach" {default}
'                 2-   "Hazen-Williams"
'    Algorithm  : as expressions Darcy-Weisbach friction factor (valid for the solver "Darcy-Weisbach)
'                 1.a- "Colebrook-White"            | Implicit
'                 1.b- "Clamond"        {default}   | Explicit
'                 1.c- "Moody"                      | Explicit
'                 1.d- "Swamee-Jain"                | Explicit
'                 1.e- "Zigrang-Sylvester"          | Explicit
'                 1.f- "Haaland"                    | Explicit
'   OUTPUT
'    PressureLoss: Pressure Loss                [bar]


Dim Re As Double, rRou As Double, vFlow As Double
Re = Reynolds(mFlow, D, T, P)       ' Calculation of Reynolds
vFlow = mFlow / rhoL_T(T)           ' Calculation of volumetric flow (rhol_T returns density at temperature T, by the XSteam module)
rRou = aRou / D                     ' Relative roughness

' Fasten your seat belts, Calculation on Progress (based on the solver and the algorithm selection by user)
Select Case Solver
    Case "Hazen-Williams"
        ' Limitations of Hazen-Williams
        If D < 50 Or D > 1850 Then                  ' Diameter Check
            MsgBox "Hazen-Williams equation is valid for diameter range as in 0.05<D<1.85 m"
            PressureLoss = CVErr(xlErrNA)
            Exit Function
            End If
        If T > 15 Then                              ' Temperature Check
            MsgBox "Hazen-Williams equation is valid at a temperature range of 4-15 °C"
            PressureLoss = CVErr(xlErrNA)
            Exit Function
            End If
        ' Reynolds Limitation
        Dim LimitArrayRe As Variant
        LimitArrayRe = tReynoldsLimits(rRou, "rRou") ' Returning the limits from via interpolation through the DiskinData
        maxRe = LimitArrayRe(1)
        minRe = LimitArrayRe(2)
        If Re < minRe Or Re > maxRe Then             ' Reynolds Number Check
            MsgBox ("Hazen-Williams is applicable only for Reynolds Numbers " & minRe & "<Re<" & maxRe & " for the given relative roughness value as " & rRou & " [mm/mm] - ref: Diskin")
            PressureLoss = CVErr(xlErrNA)
            Exit Function
            End If
        
        ' Calculation by Hazen-Williams formulation
        c = tConverterRoughness(rRou, "rRou2C")     ' Returning Hazen-Williams roughness coefficient C by converting from the relative roughness
        PressureLoss = (L * (vFlow / (0.278 * c * (D / 1000) ^ 2.63)) ^ 1.85185) * 0.09807 '[bar]
        
    Case "Darcy-Weisbach"
        If Re < 2300 Then           ' Laminar zone
            f = 64 / Re
        Else
            Select Case Algorithm   ' Transient and Turbulent zone
                Case "Colebrook-White"
                    f = f_ColebrookWhite(D, Re, aRou, fTol, MaxIter)
                Case "Moody"
                    f = f_Moody(D, Re, aRou)
                Case "Clamond"
                    f = f_Clamond(Re, rRou)
                Case "Haaland"
                    f = f_Haaland(D, Re, aRou)
                Case "Swamee-Jain"
                    f = f_SwameeJain(D, Re, aRou)
                Case "Zigrang-Sylvester"
                    f = f_ZigrangSylvester(D, Re, aRou)
                Case Else
                    MsgBox "Please define an algorithm, as either -Colebrook-White-, -Moody-, -Clamond-, -Haaland-, -Swamee-Jain-, or -Zigrang-Sylvester-"
                    PressureLoss = CVErr(xlErrNA)
                    Exit Function
            End Select
         End If
         
         PressureLoss = (8 * f * L * mFlow ^ 2 / (PiNumber() ^ 2 * rhoL_T(T) ^ 2 * 9.81 * (D / 1000) ^ 5)) / 10.1971621297792 '[bar]
                
                
    Case Else
        MsgBox "Please define a solver, as either -Hazen-Williams- or -Darcy-Weisbach-"
        PressureLoss = CVErr(xlErrNA)
        Exit Function
End Select
End Function




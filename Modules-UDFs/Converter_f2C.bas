Attribute VB_Name = "Converter_f2C"
Function tConverterDW2HW(f_or_C As Double, D As Double, Re As Double, T As Double, P As Double, ConvertDir As String) As Double
' Converts the Darcy-Weisbach friction factor (f) to Hazen-Williams pipe roughness coefficient (C)
'   formulation based on expression given by Liou
'   https://doi.org/10.1061/(ASCE)0733-9429(1998)124:9(951)
'  by Tol,Hakan Ibrahim from the PhD study at Technical University of Denmark
'   PhD Topic: District Heating in Areas with Low-Energy Houses

' NOTE
'   For PEX material, C is given as 150, ref: REHAU, Committee report (2012)
'   Here the idea is to provide a dynamically changing C as to the f (not constant C)

' INPUTS
'   f        : Darcy-Weisbach fricton factor  [-]
'   D        : Inner diameter of the pipe     [mm]
'   aRou     : Absolute roughness             [mm]
'   Re       : Reynolds Number                [-]
'   T        : Temperature                    [ºC]
'   P        : Hydraulic static pressure      [bar]

' Fasten your seat belts
'   Converting dynamic viscosity to kinematic viscosity

kVisco = my_pT(P, T) / rhoL_T(T) ' as dynamic viscosity (my_pT)/density(rhol_T) [m2/s] (XSteam functions in paranthesis)

Select Case ConvertDir
    Case "f2C"
        tConverterDW2HW = 1 / ((5 * D ^ (79 / 5000) * Re ^ (37 / 250) * f_or_C * kVisco ^ (37 / 250)) / 669) ^ (20 / 37)
    Case "C2f"
        tConverterDW2HW = 669 / (5 * f_or_C ^ (37 / 20) * D ^ (79 / 5000) * Re ^ (37 / 250) * kVisco ^ (37 / 250))
    Case Else
        MsgBox "You must choose either -f2C- or -C2f- to indicate convert direction"
End Select
End Function





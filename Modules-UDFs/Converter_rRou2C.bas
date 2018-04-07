Attribute VB_Name = "Converter_rRou2C"
Function tConverterRoughness(rRou_or_C As Double, ConvertDir As String) As Double
' BiDirectional convertor tool for pipe roughness as 'e/D to C' or 'C to e/D' - rRou is relative roughness (e/D)
'   based on data given by Diskin,M.H. - The limits of applicability of the Hazen-Williams formula
'   The data by Diskin can be found in the Excel sheet "zDiskinData"
'  by Tol,Hakan Ibrahim from the PhD study at Technical University of Denmark
'   PhD Topic: District Heating in Areas with Low-Energy Houses

' WARNING: This VBA code uses Diskin Data ranges named as "rRou_Data" and "Cmod_Data" given in sheet "zDiskinData"

' DESRIPTION
'   INPUTS
'    rRou_C     : Relative roughness (eps/D) or roughness coefficient (C), used respectively in Darcy-Weisbach (DW) or Hazen-Williams (HW) formulations
'                 Depends on the user choice of the conversion direction
'    ConvertDir : Direction for conversion, as either 'rRou2C' for 'e/D to C' or 'C2rRou for 'C to e/D'

'   OUTPUT
'    tConverterRoughness: Roughness coefficient (C) for HW or relative roughness (eps/D) for DW depending on the convert direction selected

' Calculations as to the Direction of Conversion given as input 'rRou2C' or 'C2rRou'

Select Case ConvertDir
    Case "rRou2C" ' C as a function of relative roughness rRou
        tConverterRoughness = -147.120418481659 * rRou_or_C ^ 0.168899561932556 + 179.126967041124
    Case "C2rRou" ' rRou as a function of C
        tConverterRoughness = Linterp(Range("rRou_Data"), Range("Cmod_Data"), rRou_or_C)
    Case Else
        MsgBox "You must choose either -rRou2C- or -C2rRou- to indicate convert direction"
End Select

End Function


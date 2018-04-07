Attribute VB_Name = "tReynolds"
Function Reynolds(mFlow, D, T, P)
' Calculates the Reynolds Number for a circular pipe (water)

'  Description
'  INPUTS
'   mFlow   : Mass flow                     [kg/s]
'   D       : Inner diameter of the pipe    [mm]
'   T       : Temperature of the water      [°C]

Reynolds = 4 * mFlow / (PiNumber() * (D / 1000) * my_pT(P, T))

End Function

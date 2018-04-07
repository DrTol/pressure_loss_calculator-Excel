Attribute VB_Name = "DWf_Clamond"
Function f_Clamond(Re As Double, rRou As Double) As Double
' F = COLEBROOK(R ,K) fast , accurate and robust computation of the Darcy-Weisbach friction factor according to the Colebrook formula

' Ref: Clamond D. Efficient resolution of the colebrook equation. Industrial & Engineering Chemistry Research 2009; 48: p. 3665-3671.
' Link: https://arxiv.org/pdf/0810.5564.pdf
' This VBA code is modified by Hakan ibrahim Tol from the Matlab Code by Clamond D,
' its link: https://nl.mathworks.com/matlabcentral/fileexchange/21990-colebrook-m?focused=5105324&tab=function

' INPUTS:
'   R | Re      : Reynolds' number (should be >= 2300).
'   K | rRou    : Equivalent sand roughness height divided by the hydraulic diameter (default K=0).

' RE-ARRANGING INPUT (as to Clamond Argument Names)

R = Re
k = rRou

' INPUT VALUE CHECK

If R < 2300 Then
    f_Clamond = CVErr(xlErrNA)
    MsgBox "The Colebrook equation is valid for Reynolds numbers (Re | R) >= 2300"
End If

If k < 0 Then
    f_Clamond = CVErr(xlErrNA)
    MsgBox "The relative sand roughness (rRou | K) must be non-negative"
End If

' Fasten your seat belts - Iteration Starts

'Initialization
X1 = k * R * 0.123968186335418       ' X1 <- K * R * log(10) / 18.574.
X2 = Log(R) - 0.779397488455682      ' X2 <- log( R * log(10) / 5.02 );

'Initial Guess
f = X2 - 0.2

' First Iteration
E = (Log(X1 + f) + f - X2) / (1 + X1 + f)
f = f - (1 + X1 + f + 0.5 * E) * E * (X1 + f) / (1 + X1 + f + E * (1 + E / 3))

' Second Iteration (remove the next two lines for moderate accuracy).
E = (Log(X1 + f) + f - X2) / (1 + X1 + f)
f = f - (1 + X1 + f + 0.5 * E) * E * (X1 + f) / (1 + X1 + f + E * (1 + E / 3))

' Finalized Solution                ' F <- 0.5 * log(10) / F;
f = 1.15129254649702 / f            ' F <- Friction factor.
f_Clamond = f * f
    
End Function

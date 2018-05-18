# Pricing-American-Options
##### Pricing American Options using the False Position Method
##### Apple Stock was taken, with a time step of one month for 5 periods.
##### The options were priced using implied volatility trees (IVT) which followed a Derman and Kani Approach.
##### Firstly a CRR tree was calculated, secondly a Derman and Kani IVT for Apple stock on European Options and lastly the Derman and Kani for American. A False position method was used (similiar to bisection method) to get around the problem of early exercise in American Option.

##### Below is the VBA code used for the False Position Method.
```VBA
Public Function FP(Pam As Double, sigmaEUR As Double, sigmaAMR As Double, lamda As Double)
'Declare Variables
 Dim V0 As Double
 Dim V1 As Double
 Dim m As Double
'For False position we get our guesses
'and Pvalues
'P is Pam - american price we interpolated
V0 = (Pam - sigmaEUR) / lamda
V1 = 0
PV0 = sigmaAMR + lamda * V0 - Pam
PV1 = sigmaAMR + lamda * V1 - Pam
'Incase is already close
If VBA.Round(PV0, 1) = 0 Then
     FP = V0
 Else
'now need to set up false position
     m = (PV1 - PV0) / (V1 - V0)
     V2 = (V1 * m - PV1) / m
     PV2 = sigmaAMR + lamda * V2 - Pam
     If VBA.Round(PV2, 1) = 0 Then
         FP = V2
     Else
         Do
           If PV0 * PV2 < 0 Then
              m = (PV2 - PV0) / (V2 - V0)
              V2 = (V2 * m - PV2) / m
              PV2 = sigmaAMR + lamda * V2 - Pam
                If PV0 * PV2 > 0 Then
                    PV0 = PV2
                    V0 = V2
                End If
           End If
                If PV1 * PV2 < 0 Then
                m = (PV2 - PV1) / (V2 - V1)
                V2 = (V2 * m - PV2) / m
                PV2 = sigmaAMR + lamda * V2 - Pam
                    If PV2 * PV1 > 0 Then
                        PV1 = PV2
                        V1 = V2
                    End If
                End If
           
         Loop Until VBA.Round(PV2, 1) = 0
         FP = V2
     End If
 End If
End Function
```

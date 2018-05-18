# Pricing-American-Options
##### Pricing American Options using the False Position Method
##### Apple Stock (dividened included - schedule in Excel sheet) was taken, with a time step of one month for 5 periods.
##### The options were priced using implied volatility trees (IVT) which followed a Derman and Kani Approach.
##### Firstly a CRR tree was calculated, secondly a Derman and Kani IVT for Apple stock on European Options and lastly the Derman and Kani for American. A False position method was used (similiar to bisection method) to get around the problem of early exercise in American Option.

##### Below is the VBA code used for the False Position Method:
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
##### To attain the implied volatilities for the option needed to be priced, Bi-Linear Interpolation was used, whereby using strike and maturties above and below the option being priced. The VBA code below can be used:
```VBA
Function BInt(iopt, i, K, Tp1, rngtm1 As Range, rngtp1 As Range)
't-1
indexuptm1 = 3
Kuptm1 = 0
Kdowntm1 = 0
    For Each elem In rngtm1
    indexuptm1 = indexuptm1 + 1
        If elem.Value > K Then
            Kuptm1 = elem.Value
            Exit For
        End If
        Kdowntm1 = elem.Value
    Next
    indexdowntm1 = indexuptm1 - 1

't+1
indexuptp1 = 3
Kuptp1 = 0
Kdowntp1 = 0
    For Each elem In rngtp1
    indexuptp1 = indexuptp1 + 1
        If elem.Value > K Then
            Kuptp1 = elem.Value
            Exit For
        End If
        Kdowntp1 = elem.Value
    Next
    indexdowntp1 = indexuptp1 - 1

    If iopt = 1 Then
        voldowntm1 = (Sheets(i).Cells(indexdowntm1, 6).Value) / 100
        voluptm1 = (Sheets(i).Cells(indexuptm1, 6).Value) / 100
        voldowntp1 = (Sheets(i + 2).Cells(indexdowntp1, 6).Value) / 100
        voluptp1 = (Sheets(i + 2).Cells(indexuptp1, 6).Value) / 100
    ElseIf iopt = -1 Then
        voldowntm1 = (Sheets(i).Cells(indexdowntm1, 14).Value) / 100
        voluptm1 = (Sheets(i).Cells(indexuptm1, 14).Value) / 100
        voldowntp1 = (Sheets(i + 2).Cells(indexdowntp1, 14).Value) / 100
        voluptp1 = (Sheets(i + 2).Cells(indexuptp1, 14).Value) / 100
    End If
    
    Tm1 = 0
    T = 1 / 12
    'NOTE THESE ARE THE FORMULA NEEDED FOR BI_LINEAR INTERPOLATION
    a = (T - Tm1) / (Tp1 - Tm1)
    b = (K - Kdowntm1) / (Kuptm1 - Kdowntm1)
    
    imvol = (1 - a) * (1 - b) * voldowntm1 + a * (1 - b) * voldowntp1 + a * b * voluptp1 + b * (1 - a) * voluptm1
    BInt = imvol
    
End Function
```

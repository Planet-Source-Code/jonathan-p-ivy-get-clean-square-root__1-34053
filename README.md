<div align="center">

## Get Clean Square Root


</div>

### Description

It returns a nice clean square root of a number (No Fractions). It uses a simple factoring routine to get all the factors of a number. It then multiplies all the terms that appear twice. It returns "N1~N2", where N1*sqr(N2)=sqr(X).
 
### More Info
 
X

"N1~N2", where N1*sqr(N2)=sqr(X).


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jonathan P\. Ivy](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jonathan-p-ivy.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jonathan-p-ivy-get-clean-square-root__1-34053/archive/master.zip)





### Source Code

```
Function GetSqr(ByVal X As Double) As String
Dim Y As Double, Z As Double, N1 As Double, N2 As Double
Dim NA() As String
Z = X
Y = 1
ReDim NA(0)
Do
Z = Z / Y
Y = GetLF(Z)
ReDim Preserve NA(UBound(NA) + 1)
NA(UBound(NA)) = Y
If Y = Z Or Y = 1 Then Exit Do
Loop
Debug.Print Join(NA, " ")
N1 = 1
N2 = 1
For Y = 1 To UBound(NA)
For Z = Y To UBound(NA)
If Z <> Y And NA(Z) = NA(Y) Then N1 = N1 * NA(Z): NA(Z) = 1: Exit For
Next
If Z > UBound(NA) Then N2 = N2 * NA(Y)
NA(Y) = 1
Next
If N2 > 1 Then GetSqr = N1 & "~" & N2 Else GetSqr = N1
End Function
Function GetLF(X As Double)
For N1 = 2 To Fix(X / 2) + 1
If Fix(X / N1) = (X / N1) Then GetLF = N1: Exit Function
Next
GetLF = X
End Function
```


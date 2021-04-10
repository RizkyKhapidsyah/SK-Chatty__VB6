Attribute VB_Name = "RSA64"
Public mark As Byte
Public key(1 To 3) As Double
Public p As Double, q As Double
Public PHI As Double

Public Function Mult(ByVal x As Double, ByVal pg As Double, ByVal m As Double) As Double
On Error GoTo error1
y = 1
    Do While pg > 0
        Do While (pg / 2) = Int((pg / 2))
            x = nMod((x * x), m)
            pg = pg / 2
        Loop
        y = nMod((x * y), m)
        pg = pg - 1
    Loop
    Mult = y
    Exit Function
error1:
y = 0
End Function

Private Function nMod(x As Double, y As Double) As Double
  On Error Resume Next
  Dim z#
  z = x - (Int(x / y) * y)
  nMod = z
End Function

Public Function enc(ByVal tIp As String, eE As Double, eN As Double) As String
On Error Resume Next
Dim encSt As String
encSt = ""
e2st = ""
    If tIp = "" Then Exit Function
    For I = 1 To Len(tIp)
        encSt = encSt & Mult(CLng(Asc(Mid(tIp, I, 1))), eE, eN) & "+"
    Next I
enc = encSt
End Function

Public Function dec(ByVal tIp As String, dD As Double, dN As Double) As String
On Error Resume Next
Dim decSt As String
decSt = ""
For z = 1 To Len(tIp)
    ptr = InStr(z, tIp, "+")
    tok = Val(Mid(tIp, z, ptr))
    decSt = decSt + Chr(Mult(tok, dD, dN))
    z = ptr
Next z
dec = decSt
End Function




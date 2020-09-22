Attribute VB_Name = "Module1"
Option Explicit
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Dim xf As Single, yf As Single, xi As Long, yi As Long
Dim tlColR As Long, trColR As Long, blColR As Long, brColR As Long
Dim tlColG As Long, trColG As Long, blColG As Long, brColG As Long
Dim tlColB As Long, trColB As Long, blColB As Long, brColB As Long
Dim tlCurR As Long, trCurR As Long, blCurR As Long, brCurR As Long
Dim tlCurG As Long, trCurG As Long, blCurG As Long, brCurG As Long
Dim tlCurB As Long, trCurB As Long, blCurB As Long, brCurB As Long
Dim tlR As Long, trR As Long, blR As Long, brR As Long
Dim tlG As Long, trG As Long, blG As Long, brG As Long
Dim tlB As Long, trB As Long, blB As Long, brB As Long

Sub WuPset(wx As Single, wy As Single, colour As Long)
If Sgn(wx) = -1 Or Sgn(wy) = -1 Then Exit Sub
'Split coordinates into fixed floating point numbers and integers
xi = Int(wx): yi = Int(wy)
xf = Int((wx - Fix(wx)) * 100) / 100: yf = Int((wy - Fix(wy)) * 100) / 100
'Work out colours of 2x2 pixel group
'RED
tlColR = (1 - xf) * (1 - yf) * (colour And &HFF)
trColR = (xf) * (1 - yf) * (colour And &HFF)
blColR = (1 - xf) * (yf) * (colour And &HFF)
brColR = (xf) * (yf) * (colour And &HFF)
'GREEN
tlColG = (1 - xf) * (1 - yf) * (colour And &HFF00&) / &H100&
trColG = (xf) * (1 - yf) * (colour And &HFF00&) / &H100&
blColG = (1 - xf) * (yf) * (colour And &HFF00&) / &H100&
brColG = (xf) * (yf) * (colour And &HFF00&) / &H100&
'BLUE
tlColB = (1 - xf) * (1 - yf) * (colour And &HFF0000) / &H10000
trColB = (xf) * (1 - yf) * (colour And &HFF0000) / &H10000
blColB = (1 - xf) * (yf) * (colour And &HFF0000) / &H10000
brColB = (xf) * (yf) * (colour And &HFF0000) / &H10000
'Retreive the current RGB value
'RED
tlCurR = GetPixel(Form1.hdc, xi, yi) And &HFF
trCurR = GetPixel(Form1.hdc, xi + 1, yi) And &HFF
blCurR = GetPixel(Form1.hdc, xi, yi + 1) And &HFF
brCurR = GetPixel(Form1.hdc, xi + 1, yi + 1) And &HFF
'GREEN
tlCurG = (GetPixel(Form1.hdc, xi, yi) And &HFF00&) / &H100&
trCurG = (GetPixel(Form1.hdc, xi + 1, yi) And &HFF00&) / &H100&
blCurG = (GetPixel(Form1.hdc, xi, yi + 1) And &HFF00&) / &H100&
brCurG = (GetPixel(Form1.hdc, xi + 1, yi + 1) And &HFF00&) / &H100&
'BLUE
tlCurB = (GetPixel(Form1.hdc, xi, yi) And &HFF0000) / &H10000
trCurB = (GetPixel(Form1.hdc, xi + 1, yi) And &HFF0000) / &H10000
blCurB = (GetPixel(Form1.hdc, xi, yi + 1) And &HFF0000) / &H10000
brCurB = (GetPixel(Form1.hdc, xi + 1, yi + 1) And &HFF0000) / &H10000

If tlCurR > tlColR Then tlR = tlCurR Else tlR = tlColR
If tlCurG > tlColG Then tlG = tlCurG Else tlG = tlColG
If tlCurB > tlColB Then tlB = tlCurB Else tlB = tlColB

If trCurR > trColR Then trR = trCurR Else trR = trColR
If trCurG > trColG Then trG = trCurG Else trG = trColG
If trCurB > trColB Then trB = trCurB Else trB = trColB

If blCurR > blColR Then blR = blCurR Else blR = blColR
If blCurG > blColG Then blG = blCurG Else blG = blColG
If blCurB > blColB Then blB = blCurB Else blB = blColB

If brCurR > brColR Then brR = brCurR Else brR = brColR
If brCurG > brColG Then brG = brCurG Else brG = brColG
If brCurB > brColB Then brB = brCurB Else brB = brColB

'Finally, plot the pixels

SetPixel Form1.hdc, xi, yi, RGB(tlR, tlG, tlB)
SetPixel Form1.hdc, xi + 1, yi, RGB(trR, trG, trB)
SetPixel Form1.hdc, xi, yi + 1, RGB(blR, blG, blB)
SetPixel Form1.hdc, xi + 1, yi + 1, RGB(brR, brG, brB)
End Sub


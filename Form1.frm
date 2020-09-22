VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Wu Pixel Demo"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4770
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   132
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   $"Form1.frx":0000
      ForeColor       =   &H00C0FFFF&
      Height          =   1755
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const PI As Single = 3.141592654

Dim a As Long, n As Single, xx As Single, yy As Single

Private Sub Form_Activate()
Form1.Refresh
For a = 0 To 360 Step 1
n = a * (PI / 180)
xx = CSng(Sin(n) * 1)
yy = CSng(Cos(n) * 1)
PSet (xx + 30, yy + 10), vbWhite
WuPset xx + 70, yy + 10, vbWhite
Next
Form1.Refresh

For a = 0 To 360 Step 1
n = a * (PI / 180)
xx = CSng(Sin(n) * 2)
yy = CSng(Cos(n) * 2)
PSet (xx + 30, yy + 20), vbWhite
WuPset xx + 70, yy + 20, vbWhite
Next
Form1.Refresh

For a = 0 To 360 Step 1
n = a * (PI / 180)
xx = CSng(Sin(n) * 5)
yy = CSng(Cos(n) * 5)
PSet (xx + 30, yy + 35), vbWhite
WuPset xx + 70, yy + 35, vbWhite
Next
Form1.Refresh

For a = 0 To 360 Step 1
n = a * (PI / 180)
xx = CSng(Sin(n) * 10)
yy = CSng(Cos(n) * 10)
PSet (xx + 30, yy + 60), vbWhite
WuPset xx + 70, yy + 60, vbWhite
Next
Form1.Refresh

For a = 0 To 360 Step 1
n = a * (PI / 180)
xx = CSng(Sin(n) * 15)
yy = CSng(Cos(n) * 15)
PSet (xx + 30, yy + 95), vbWhite
WuPset xx + 70, yy + 95, vbWhite
Next
Form1.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

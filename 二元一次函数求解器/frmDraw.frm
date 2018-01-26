VERSION 5.00
Begin VB.Form frmDraw 
   AutoRedraw      =   -1  'True
   Caption         =   "º¯ÊýÍ¼Ïó"
   ClientHeight    =   4290
   ClientLeft      =   255
   ClientTop       =   555
   ClientWidth     =   7110
   Icon            =   "frmDraw.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7110
End
Attribute VB_Name = "frmDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
frmDraw.Scale (-10, 25)-(10, -25)
frmDraw.Line (-10, 0)-(10, 0), RGB(0, 0, 255)
frmDraw.Line (0, 25)-(0, -25), RGB(0, 0, 255)
For x = -10 To 10 Step 0.0001
    y = a * (x ^ 2) + b * x + c
    frmDraw.PSet (x, y), RGB(0, 255, 0)
Next x
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmResult
frmMain.Show
End Sub

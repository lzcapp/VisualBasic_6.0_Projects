VERSION 5.00
Begin VB.Form frmAbouta 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7200
      Begin VB.Timer Timer1 
         Left            =   120
         Top             =   960
      End
      Begin VB.Label Label2 
         Caption         =   "作者： 罗子琛"
         BeginProperty Font 
            Name            =   "隶书"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   3480
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "在此向罗子琛父母亲致以最崇高的敬意！"
         BeginProperty Font 
            Name            =   "华文行楷"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   480
         TabIndex        =   3
         Top             =   1320
         Width           =   6375
      End
      Begin VB.Label lblCopyright 
         Caption         =   "本程序版权受中华人民共和国及国际法律、条约保护。                                      侵权必究！"
         BeginProperty Font 
            Name            =   "隶书"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   450
         TabIndex        =   1
         Top             =   2640
         Width           =   6495
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "二次函数求解器"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   42
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   743
         TabIndex        =   2
         Top             =   240
         Width           =   5880
      End
   End
End
Attribute VB_Name = "frmAbouta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim t

Private Sub Form_KeyPress(KeyAscii As Integer)
    frmMain.Show
    Unload Me
End Sub

Private Sub Form_Load()
    Timer1.Interval = 1000
End Sub

Private Sub Frame1_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Timer1_Timer()
    t = t + 1
    If t = 3 Then frmMain.Show
    If t = 3 Then Unload Me
End Sub

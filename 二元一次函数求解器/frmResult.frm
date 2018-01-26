VERSION 5.00
Begin VB.Form frmResult 
   Caption         =   "一元二次函数分析结果"
   ClientHeight    =   5790
   ClientLeft      =   6750
   ClientTop       =   4770
   ClientWidth     =   8415
   Icon            =   "frmResult.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8415
   Begin VB.Label Labely 
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   37
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label17 
      Caption         =   ","
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   36
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Labelx 
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   35
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label15 
      Caption         =   "("
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   34
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "对称轴"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   33
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "直线 x ="
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   32
      Top             =   2160
      Width           =   1560
   End
   Begin VB.Label Labeldc 
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   31
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Label Label11 
      Caption         =   "顶点坐标"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   30
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "与x轴交点"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "与y轴交点"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label21 
      Caption         =   "("
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   27
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Labeljxo 
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   26
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label23 
      Caption         =   ","
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   25
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label24 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   24
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label25 
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   23
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label26 
      Caption         =   "("
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   22
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label27 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   21
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label28 
      Caption         =   ","
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   20
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Labeljy 
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   19
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "("
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   18
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Labeljxt 
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   ","
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   16
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label16 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "开口方向"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Labelfx 
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label20 
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   12
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label30 
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   11
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label18 
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Top             =   5160
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   2520
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   2520
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   2760
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   2760
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   5520
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   5760
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   5520
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape8 
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   5760
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Labelt 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   165
      TabIndex        =   9
      Top             =   840
      Width           =   8055
   End
   Begin VB.Label Label8 
      Caption         =   "结果分析区"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3285
      TabIndex        =   8
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   900
      TabIndex        =   5
      Top             =   1500
      Width           =   375
   End
   Begin VB.Label Labelxo 
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1455
      TabIndex        =   4
      Top             =   1560
      Width           =   2520
   End
   Begin VB.Label Label13 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4740
      TabIndex        =   2
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label14 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5100
      TabIndex        =   1
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Labelxt 
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5655
      TabIndex        =   0
      Top             =   1560
      Width           =   2520
   End
   Begin VB.Label Label19 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   255
      TabIndex        =   7
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4500
      TabIndex        =   3
      Top             =   1500
      Width           =   495
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Labelxo.Caption = ""
Labelxt.Caption = ""
Labelx.Caption = ""
Labely.Caption = ""
Labeldc.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmDraw
frmMain.Texta.Text = ""
frmMain.Textb.Text = ""
frmMain.Textc.Text = ""
frmMain.Show
End Sub


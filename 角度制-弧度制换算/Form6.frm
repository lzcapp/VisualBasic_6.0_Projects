VERSION 5.00
Object = "{F6B9D3BC-3953-4A68-AD1A-BD05206D76A9}#1.0#0"; "hmButton.ocx"
Begin VB.Form frmYuan1 
   Caption         =   "原理"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7050
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   7050
   StartUpPosition =   2  '屏幕中心
   Begin 黑马按钮控件.dcButton dcButton1 
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   6720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      BackColor       =   14995922
      ButtonStyle     =   9
      Caption         =   "返回"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   3360
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line2 
      X1              =   3600
      X2              =   3600
      Y1              =   2640
      Y2              =   960
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "公式："
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "程序框图："
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Line Line1 
      X1              =   3840
      X2              =   3840
      Y1              =   2280
      Y2              =   7680
   End
   Begin VB.Image Image2 
      Height          =   1290
      Left            =   360
      Picture         =   "Form6.frx":0000
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   3225
   End
   Begin VB.Image Image1 
      Height          =   7695
      Left            =   3960
      Picture         =   "Form6.frx":0785
      Stretch         =   -1  'True
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "frmYuan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: 使用 Option Explicit 可以避免隐式创建 Variant 类型的变量                                          FixIT90210ae-R383-H1984
Private Sub dcButton1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.BackColor = RGB(187, 223, 187)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form1.Show
Me.Hide
End Sub


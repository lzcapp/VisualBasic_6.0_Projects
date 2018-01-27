VERSION 5.00
Object = "{F6B9D3BC-3953-4A68-AD1A-BD05206D76A9}#1.0#0"; "hmButton.ocx"
Begin VB.Form frmIn1 
   Caption         =   "输入"
   ClientHeight    =   2310
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   3900
   StartUpPosition =   2  '屏幕中心
   Begin 黑马按钮控件.dcButton dcButton1 
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BackColor       =   1228031
      ButtonStyle     =   11
      Caption         =   "确定"
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3375
   End
   Begin 黑马按钮控件.dcButton dcButton2 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BackColor       =   1228031
      ButtonStyle     =   11
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "请输入一个角度值："
      BeginProperty Font 
         Name            =   "华文仿宋"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   720
      Width           =   255
   End
End
Attribute VB_Name = "frmIn1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: 使用 Option Explicit 可以避免隐式创建 Variant 类型的变量                                          FixIT90210ae-R383-H1984

Private Sub dcButton1_Click()
If IsNumeric(Text1.Text) Then
dus = Text1.Text
Me.Hide
frmRe1.Show

Else
a = MsgBox("请输入数字！", , "错误")
End If
End Sub

Private Sub dcButton2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.BackColor = RGB(187, 223, 187)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form1.Show
Me.Hide
End Sub

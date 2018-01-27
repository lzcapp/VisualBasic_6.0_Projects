VERSION 5.00
Object = "{F6B9D3BC-3953-4A68-AD1A-BD05206D76A9}#1.0#0"; "hmButton.ocx"
Begin VB.Form frmRe1 
   Caption         =   "结果"
   ClientHeight    =   2460
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4560
   StartUpPosition =   2  '屏幕中心
   Begin 黑马按钮控件.dcButton dcButton1 
      Height          =   495
      Left            =   413
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BackColor       =   14995922
      ButtonShape     =   3
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
   Begin 黑马按钮控件.dcButton dcButton2 
      Height          =   495
      Left            =   2693
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BackColor       =   14995922
      ButtonShape     =   3
      ButtonStyle     =   9
      Caption         =   "原理"
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
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "π"
      BeginProperty Font 
         Name            =   "Gulim"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2880
      TabIndex        =   7
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "→"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Gulim"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "rad"
      BeginProperty Font 
         Name            =   "Gulim"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "frmRe1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: 使用 Option Explicit 可以避免隐式创建 Variant 类型的变量                                          FixIT90210ae-R383-H1984
'FIXIT: 用早期绑定的数据类型声明 "hu"                                                                  FixIT90210ae-R1672-R1B8ZE
Dim hu

Private Sub dcButton1_Click()
Unload Me
End Sub

Private Sub dcButton2_Click()
frmYuan1.Show
Me.Hide
End Sub

Private Sub Form_Load()
hu = dus / 180
Label1.Caption = hu
Label3.Caption = dus

Me.BackColor = RGB(187, 223, 187)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmIn1.Show
Me.Hide
End Sub


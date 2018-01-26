VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "二次函数求解器"
   ClientHeight    =   2895
   ClientLeft      =   3780
   ClientTop       =   5625
   ClientWidth     =   7860
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   7860
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4924
      TabIndex        =   12
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   885
      TabIndex        =   10
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox Textc 
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   5970
      TabIndex        =   9
      Text            =   "c"
      ToolTipText     =   "请输入数字"
      Top             =   840
      Width           =   720
   End
   Begin VB.TextBox Textb 
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   4170
      TabIndex        =   8
      Text            =   "b"
      ToolTipText     =   "请输入数字"
      Top             =   840
      Width           =   720
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   2130
      TabIndex        =   7
      Text            =   "a"
      ToolTipText     =   "请输入数字"
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Label8 
      Caption         =   "请输入需分析的二次函数一般式"
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
      Left            =   1403
      TabIndex        =   11
      Top             =   240
      Width           =   5055
   End
   Begin VB.Shape Shape8 
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   7440
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   7200
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   7440
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   7200
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   480
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   480
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   240
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   240
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5490
      TabIndex        =   6
      Top             =   960
      Width           =   345
   End
   Begin VB.Label Label6 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5010
      TabIndex        =   5
      Top             =   960
      Width           =   345
   End
   Begin VB.Label Label5 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3690
      TabIndex        =   4
      Top             =   960
      Width           =   345
   End
   Begin VB.Label Label4 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3330
      TabIndex        =   3
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2970
      TabIndex        =   2
      Top             =   960
      Width           =   345
   End
   Begin VB.Label Label2 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1650
      TabIndex        =   1
      Top             =   960
      Width           =   345
   End
   Begin VB.Label Label1 
      Caption         =   "y"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1170
      TabIndex        =   0
      Top             =   960
      Width           =   345
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If IsNumeric(Texta.Text) And IsNumeric(Textb.Text) And IsNumeric(Textc.Text) Then

a = Texta.Text
b = Textb.Text
c = Textc.Text

t = b ^ 2 - 4 * a * c
If t < 0 Then frmResult.Labelt.Caption = "该二元一次函数没有实数根"
If t < 0 Then frmResult.Labelt.ForeColor = vbRed


If t = 0 Then frmResult.Labelt.Caption = "该二元一次函数有两个相等的实数根"
If t > 0 Then frmResult.Labelt.Caption = "该二元一次函数有两个不相等的实数根"
If t >= 0 Then frmResult.Labelt.ForeColor = vbGreen

If t >= 0 Then xo = ((-b) + Sqr(b ^ 2 - 4 * a * c)) / (2 * a)
If t >= 0 Then xt = ((-b) - Sqr(b ^ 2 - 4 * a * c)) / (2 * a)

If a > 0 Then frmResult.Labelfx.Caption = "向上"
If a < 0 Then frmResult.Labelfx.Caption = "向下"

frmResult.Labelxo.Caption = xo
frmResult.Labelxo.ToolTipText = xo
frmResult.Labelxt.Caption = xt
frmResult.Labelxt.ToolTipText = xt

frmResult.Labeljxo.Caption = xo
frmResult.Labeljxo.ToolTipText = xo
frmResult.Labeljxt.Caption = xt
frmResult.Labeljxt.ToolTipText = xt

frmResult.Labeljy.Caption = c
frmResult.Labeljy.ToolTipText = c


x = b / (-2 * a)
y = (4 * a * c - b ^ 2) / (4 * a)

frmResult.Labelx.Caption = x
frmResult.Labelx.ToolTipText = x
frmResult.Labely.Caption = y
frmResult.Labely.ToolTipText = y

frmResult.Labeldc.Caption = x
frmResult.Labeldc.ToolTipText = x

frmResult.Show
frmDraw.Show
frmMain.Visible = False

Else
frmMain.Enabled = False
frmResult.Enabled = False
frmResult.Visible = False
frmErr.Show
Texta.Text = ""
Textb.Text = ""
Textc.Text = ""

End If

End Sub

Private Sub Command2_Click()
frmMain.Enabled = False
frmAbout.Show
End Sub

Private Sub Command3_Click()
frmBy.Show
End Sub

Private Sub Form_Load()
frmResult.Labelxo.Caption = ""
frmResult.Labelxt.Caption = ""
frmResult.Labelx.Caption = ""
frmResult.Labely.Caption = ""
frmResult.Labeldc.Caption = ""
frmMain.Caption = "二次函数求解器     by 罗子琛"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmBy.Show
End Sub

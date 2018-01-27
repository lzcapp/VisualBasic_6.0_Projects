VERSION 5.00
Object = "{F6B9D3BC-3953-4A68-AD1A-BD05206D76A9}#1.0#0"; "hmButton.ocx"
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "SysTray.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2415
   ClientLeft      =   5325
   ClientTop       =   1470
   ClientWidth     =   4725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin SysTrayCtl.cSysTray cSysTray1 
      Left            =   1080
      Top             =   600
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   -1  'True
      TrayIcon        =   "Form1.frx":0000
      TrayTip         =   "您还没有开始计时"
   End
   Begin 黑马按钮控件.dcButton dcButton3 
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BackColor       =   0
      ButtonStyle     =   4
      Caption         =   "×"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16776960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   360
   End
   Begin 黑马按钮控件.dcButton dcButton1 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      BackColor       =   16776960
      ButtonShape     =   1
      ButtonStyle     =   4
      Caption         =   "开始计时"
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
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      BackColor       =   16776960
      ButtonShape     =   1
      ButtonStyle     =   4
      Caption         =   "清零"
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
   Begin 黑马按钮控件.dcButton dcButton4 
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BackColor       =   0
      ButtonStyle     =   4
      Caption         =   "﹣ "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16776960
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      Height          =   855
      Left            =   240
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   855
      Left            =   3240
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      X1              =   3120
      X2              =   3120
      Y1              =   360
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      X1              =   1560
      X2              =   1560
      Y1              =   360
      Y2              =   1200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   855
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Menu popup 
      Caption         =   "计时程序"
      Enabled         =   0   'False
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
      Begin VB.Menu start 
         Caption         =   "开始计时"
      End
      Begin VB.Menu stop 
         Caption         =   "清零"
      End
      Begin VB.Menu popdown 
         Caption         =   "关闭菜单"
      End
      Begin VB.Menu exit 
         Caption         =   "退出"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cSysTray1_MouseDblClick(Button As Integer, Id As Long)
Me.Show
End Sub

Private Sub cSysTray1_MouseDown(Button As Integer, Id As Long)
If Button = 2 Then
PopupMenu popup
End If
End Sub

Private Sub dcButton1_Click()
Timer1.Enabled = True
End Sub

Private Sub dcButton2_Click()
Label1.Caption = 0
Label2.Caption = 0
Label3.Caption = 0
Timer1.Enabled = False
cSysTray1.TrayTip = "您还没有开始计时"
End Sub

Private Sub dcButton3_Click()
Unload Me
End Sub

Private Sub dcButton4_Click()
cSysTray1.InTray = True
Me.Hide
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.BorderStyle = 0
Me.Caption = ""
Me.Move 0, 0, 4760, 2680
popup.Enabled = False
popup.Visible = False
End Sub

Private Sub start_Click()
Timer1.Enabled = True
End Sub

Private Sub stop_Click()
Label1.Caption = 0
Label2.Caption = 0
Label3.Caption = 0
Timer1.Enabled = False
cSysTray1.TrayTip = "您还没有开始计时"
End Sub

Private Sub Timer1_Timer()
Label3.Caption = Label3.Caption + 1
If Label3.Caption = 60 Then
Label3.Caption = 0
Label2.Caption = Label2.Caption + 1
End If
If Label2.Caption = 60 Then
Label2.Caption = 0
Label1.Caption = Label1.Caption + 1
End If
cSysTray1.TrayTip = "已过" + Label1.Caption + "时" + Label2.Caption + "分" + Label3.Caption + "秒"
End Sub

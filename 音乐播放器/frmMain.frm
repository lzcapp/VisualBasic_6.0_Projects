VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CD播放器"
   ClientHeight    =   1635
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   5985
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   5280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Left            =   4800
      Top             =   0
   End
   Begin MSComctlLib.Slider sldSlider 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin MCI.MMControl mmcMedia 
      Height          =   420
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   741
      _Version        =   393216
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      BackEnabled     =   -1  'True
      StepEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      Silent          =   -1  'True
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      Shareable       =   -1  'True
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Menu menuOpen 
      Caption         =   "打开"
      NegotiatePosition=   3  'Right
   End
   Begin VB.Menu menuClose 
      Caption         =   "关闭"
   End
   Begin VB.Menu menuExit 
      Caption         =   "退出"
   End
   Begin VB.Menu menuShow 
      Caption         =   "显示歌词"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    Dim i As Integer
    
    '随机化
    Randomize
    
    
    '初始化Multimedia MCI控件
    mmcMedia.Notify = True
    mmcMedia.Shareable = False
    mmcMedia.TimeFormat = 0
End Sub

'停止播放
Private Sub menuClose_Click()
    mmcMedia_StopClick 0
End Sub

Private Sub menuExit_Click()
    Unload Me
End Sub

'打开播放文件
Private Sub menuOpen_Click()
    '初始化CommonDialog控件
    dlgCommonDialog.FileName = ""
    dlgCommonDialog.Filter = "(*.*)|*.*|(*.CDA)|*.CDA|(*.WAV)|*.WAV|(*.MP3)|*.MP3"
    dlgCommonDialog.DialogTitle = "打开多媒体文件"
    dlgCommonDialog.Action = 1
    
    '打开要播放的曲目
    '如果没有选则播放文件
    If dlgCommonDialog.FileName = "" Then
        MsgBox "请选择要播放的文件！"
    
    '如果选则了播放文件
    Else
        '初始化Multimedia MCI控件控件
        mmcMedia.Command = "Close"
        mmcMedia.Notify = False
        mmcMedia.Wait = True
        
        '打开要播放的曲目
        mmcMedia.FileName = dlgCommonDialog.FileName
        mmcMedia.Command = "open"
        mmcMedia.Notify = True
        mmcMedia.Wait = False
        mmcMedia.Command = "play"
    
        '显示播放文件名
        If Len(dlgCommonDialog.FileName) > 30 Then
            frmMain.Caption = Left(dlgCommonDialog.FileName, 12) & _
                        "..." & Right(dlgCommonDialog.FileName, 12)
        Else
            frmMain.Caption = dlgCommonDialog.FileName
        End If
         
        '设置划块值的范围
        sldSlider.Min = 0
        sldSlider.Max = mmcMedia.Length
        
    End If
End Sub

'回放
Private Sub mmcMedia_BackClick(Cancel As Integer)
    mmcMedia.Command = "Back"
End Sub

'快进
Private Sub mmcMedia_NextClick(Cancel As Integer)
    mmcMedia.Command = "Next"
End Sub

'暂停播放
Private Sub mmcMedia_PauseClick(Cancel As Integer)
    Dim i As Integer
    
    '停止播放
    mmcMedia.Command = "Pause"
    mmcMedia.Command = "Stop"
End Sub

'启动播放
Private Sub mmcMedia_PlayClick(Cancel As Integer)
    '启动播放
    mmcMedia.Command = "Play"
End Sub

'倒带
Private Sub mmcMedia_PrevClick(Cancel As Integer)
    mmcMedia.Command = "Prev"
End Sub

'步进
Private Sub mmcMedia_StepClick(Cancel As Integer)
    mmcMedia.Command = "Step"
End Sub

'停止播放
Private Sub mmcMedia_StopClick(Cancel As Integer)
    Dim i As Integer
    
    '停止播放
    mmcMedia.Command = "Stop"
    '倒带
    mmcMedia.Command = "Prev"
End Sub


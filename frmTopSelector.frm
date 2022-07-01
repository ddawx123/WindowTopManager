VERSION 5.00
Begin VB.Form frmTopSelector 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "任意窗体置顶控制器"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5745
   Icon            =   "frmTopSelector.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdNoTopmost 
      Caption         =   "撤销置顶"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdTopmost 
      Caption         =   "置顶"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtResult 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmTopSelector.frx":000C
      Top             =   840
      Width           =   5535
   End
   Begin VB.Timer mTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   360
      Top             =   240
   End
   Begin VB.CommandButton cmdGetHwnd 
      Caption         =   "开始捕获句柄"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Copyright 2012-2022 DingStudio Technology All Rights Reserved"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   5490
   End
   Begin VB.Label lblTip 
      Caption         =   "请注意：本程序需要与被操作程序处于同一权限层级或高于该程序。否则本程序将无法取得目标程序的句柄！"
      ForeColor       =   &H00FF0000&
      Height          =   465
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   5340
   End
   Begin VB.Menu logRightBtnMenu 
      Caption         =   "日志显示区域右键菜单"
      Visible         =   0   'False
      Begin VB.Menu cleanExecuteLog 
         Caption         =   "清除执行日志"
      End
   End
End
Attribute VB_Name = "frmTopSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private rHwnd As Long
Private rTitle As String

Private Sub cleanExecuteLog_Click()
    txtResult.Text = ""
End Sub

Private Sub cmdNoTopmost_Click()
    SetWindowPos rHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub cmdTopmost_Click()
    SetWindowPos rHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    writeLog "设置置顶，被操作窗口：" & rTitle
End Sub

Private Sub Form_Activate()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    writeLog "撤销置顶，被操作窗口：" & rTitle
End Sub

Private Sub cmdGetHwnd_Click()
MsgBox "请关闭此提示后激活需要操作的窗体（任意程序均可），句柄捕捉成功后将在日志输出区域显示提示信息。", vbSystemModal + vbInformation + vbOKOnly, "操作指引信息"
cmdGetHwnd.Enabled = False
mTimer.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("确认退出吗？退出后已捕获的窗口句柄将丢失，但对该窗口的操作将维持现状哦。", vbSystemModal + vbQuestion + vbYesNo, "") = vbNo Then Cancel = -1
End Sub

Private Sub mTimer_Timer()
rHwnd = GetForegroundWindow()
If rHwnd <> Me.hwnd Then
    cmdGetHwnd.Enabled = True
    cmdTopmost.Enabled = True
    cmdNoTopmost.Enabled = True
    Dim orTitle As String * 255
    Dim rLength As Integer
    rLength = Len(orTitle) - 1
    GetWindowText rHwnd, orTitle, rLength
    rTitle = orTitle
    writeLog "捕捉成功，目标窗口名：" & rTitle
    mTimer.Enabled = False
End If
End Sub

Sub writeLog(logString As String)
txtResult.Text = txtResult & vbCrLf & logString
txtResult.SelStart = Len(txtResult.Text) '重算信息长度，自动滚动scrollbar
End Sub

Private Sub txtResult_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    txtResult.Enabled = False
    PopupMenu logRightBtnMenu
    txtResult.Enabled = True
End If
End Sub

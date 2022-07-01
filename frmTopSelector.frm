VERSION 5.00
Begin VB.Form frmTopSelector 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ⴰ���ö�������"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdNoTopmost 
      Caption         =   "�����ö�"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdTopmost 
      Caption         =   "�ö�"
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
      Caption         =   "��ʼ������"
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
      Caption         =   "��ע�⣺��������Ҫ�뱻����������ͬһȨ�޲㼶����ڸó��򡣷��򱾳����޷�ȡ��Ŀ�����ľ����"
      ForeColor       =   &H00FF0000&
      Height          =   465
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   5340
   End
   Begin VB.Menu logRightBtnMenu 
      Caption         =   "��־��ʾ�����Ҽ��˵�"
      Visible         =   0   'False
      Begin VB.Menu cleanExecuteLog 
         Caption         =   "���ִ����־"
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
    writeLog "�����ö������������ڣ�" & rTitle
End Sub

Private Sub Form_Activate()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    writeLog "�����ö������������ڣ�" & rTitle
End Sub

Private Sub cmdGetHwnd_Click()
MsgBox "��رմ���ʾ�󼤻���Ҫ�����Ĵ��壨���������ɣ��������׽�ɹ�������־���������ʾ��ʾ��Ϣ��", vbSystemModal + vbInformation + vbOKOnly, "����ָ����Ϣ"
cmdGetHwnd.Enabled = False
mTimer.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("ȷ���˳����˳����Ѳ���Ĵ��ھ������ʧ�����Ըô��ڵĲ�����ά����״Ŷ��", vbSystemModal + vbQuestion + vbYesNo, "") = vbNo Then Cancel = -1
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
    writeLog "��׽�ɹ���Ŀ�괰������" & rTitle
    mTimer.Enabled = False
End If
End Sub

Sub writeLog(logString As String)
txtResult.Text = txtResult & vbCrLf & logString
txtResult.SelStart = Len(txtResult.Text) '������Ϣ���ȣ��Զ�����scrollbar
End Sub

Private Sub txtResult_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    txtResult.Enabled = False
    PopupMenu logRightBtnMenu
    txtResult.Enabled = True
End If
End Sub

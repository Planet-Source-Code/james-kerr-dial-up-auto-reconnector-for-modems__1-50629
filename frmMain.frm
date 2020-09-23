VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DUN Reconnector"
   ClientHeight    =   3975
   ClientLeft      =   3420
   ClientTop       =   3240
   ClientWidth     =   6210
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6210
   Tag             =   "FMAIN"
   Begin VB.PictureBox picBanner 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   0
      Picture         =   "frmMain.frx":058A
      ScaleHeight     =   825
      ScaleWidth      =   6210
      TabIndex        =   6
      Top             =   0
      Width           =   6210
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "for Modem Connections"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   195
         Left            =   990
         TabIndex        =   7
         Top             =   585
         Width           =   4290
      End
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Start"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   420
      Left            =   1350
      TabIndex        =   0
      Top             =   3465
      Width           =   1185
   End
   Begin VB.Timer timCheck 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   495
      Top             =   5850
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "Options"
      Height          =   420
      Left            =   90
      TabIndex        =   1
      Top             =   3465
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Log ..."
      Height          =   2535
      Left            =   45
      TabIndex        =   4
      Top             =   855
      Width           =   6090
      Begin VB.TextBox txtLog 
         Height          =   2220
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   225
         Width           =   5910
      End
   End
   Begin VB.CommandButton cmdMin 
      Caption         =   "Send to Tray"
      Height          =   420
      Left            =   3690
      TabIndex        =   2
      Top             =   3465
      Width           =   1185
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   420
      Left            =   4950
      TabIndex        =   3
      Top             =   3465
      Width           =   1185
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function InternetGetConnectedState Lib "wininet" (lpdwFlags As Long, ByVal dwReserved As Long) As Boolean
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function InternetDial Lib "wininet.dll" (ByVal hwndParent As Long, _
                                ByVal lpszConnectoid As String, _
                                ByVal dwFlags As Long, _
                                lpdwConnection As Long, _
                                ByVal dwReserved As Long) As Long

Private Const INTERNET_CONNECTION_MODEM = 1

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const HWND_TOPMOST = -1

Private nid As NOTIFYICONDATA

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private bConnected As Boolean
Private iAttempts As Integer

Private Function GetTimeStamp() As String

    GetTimeStamp = " " & Format$(Now, "Medium Date") & " at " & Format$(Now, "Long Time")
    
End Function

Private Sub cmdExit_Click()
    
    Dim lRet As Long
    
    If cmdGo.Caption = "Stop" Then
        lRet = MsgBox("The DUN Reconnector is running - are you sure you want to exit?", vbQuestion + vbYesNo)
        If lRet = vbNo Then
            Exit Sub
        End If
    End If
    
    Unload Me
    
End Sub

Private Sub cmdGo_Click()
    
    If cmdGo.Caption = "Start" Then
        txtLog = txtLog & "Started at " & Format$(Now, "hh:mm:ss") & vbCrLf
        timCheck.Enabled = True
        cmdGo.Caption = "Stop"
    Else
        timCheck.Enabled = False
        cmdGo.Caption = "Start"
    End If
    
End Sub

Private Sub cmdMin_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub cmdOpt_Click()
    frmOpt.Show vbModal, Me
    If LenB(strISP) <> 0 Then
        cmdGo.Enabled = True
        txtLog = vbNullString
    End If
End Sub

Private Sub Form_Load()
    
    Call RestoreFormPosition(Me)
    
    If LenB(strISP) <> 0 Then
        cmdGo.Enabled = True
    Else
        txtLog = "*Select 'Options' to setup the Program*" & vbCrLf
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lRet As Long
    lRet = X / Screen.TwipsPerPixelX
    Select Case lRet
    Case WM_LBUTTONDOWN
        Me.WindowState = 0
        Me.Show
        Form_Resize
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Call SaveFormPosition(Me)

End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then
        Me.Hide
        Me.Refresh
        With nid
            .cbSize = Len(nid)
            .hwnd = Me.hwnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallBackMessage = WM_MOUSEMOVE
            .hIcon = Me.Icon
            .szTip = Me.Caption & vbNullChar
        End With
        Shell_NotifyIcon NIM_ADD, nid
    Else
        Shell_NotifyIcon NIM_DELETE, nid
        Me.Show
        Me.Refresh
    End If

End Sub

Private Sub timCheck_Timer()
    
    Dim lRet As Long
    Dim bRet As Boolean

    bRet = InternetGetConnectedState(lRet, 0)
    If bRet Then
        If lRet And INTERNET_CONNECTION_MODEM Then
            bConnected = True
            Exit Sub
        End If
    End If
    
    txtLog = txtLog & "Disconnected at" & GetTimeStamp() & vbCrLf
    txtLog = txtLog & "  ... waiting for " & Format$(lWaitBeforeReconnect, "0") & " minutes" & vbCrLf
    Call Sleep(CLng((lWaitBeforeReconnect * 1000) * 60))
    
    timCheck.Enabled = False
    bConnected = False
    txtLog = txtLog & "  ... attempting reconnection to " & strISP & " :" & GetTimeStamp() & vbCrLf

GetConn:
    lRet = InternetDial(Me.hwnd, strISP, 2, 0, 0&)
    DoEvents
    Call Sleep(10000)
    bRet = InternetGetConnectedState(lRet, 0)
    If bRet Then
        If lRet And INTERNET_CONNECTION_MODEM Then
            bConnected = True
            iAttempts = 0
            bConnected = True
            txtLog = txtLog & "  ... reconnected :" & GetTimeStamp() & vbCrLf
            timCheck.Enabled = True
            Exit Sub
        End If
    End If
    
    iAttempts = iAttempts + 1
    If iAttempts > 5 Then
        txtLog = txtLog & "  ... reconnection failed :" & GetTimeStamp() & vbCrLf
    Else
        GoTo GetConn
    End If

End Sub

Private Sub txtLog_Change()
    
    If Len(txtLog) > 10000 Then
        txtLog = vbNullString
        txtLog = txtLog & "Logfile reset :" & GetTimeStamp() & vbCrLf
    End If
    txtLog.SelStart = Len(txtLog)
        
End Sub

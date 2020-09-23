VERSION 5.00
Begin VB.Form frmOpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   2565
   ClientLeft      =   3810
   ClientTop       =   5685
   ClientWidth     =   6465
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Tag             =   "FOPT"
   Begin VB.Frame Frame1 
      Caption         =   "Program Options ..."
      Height          =   1905
      Left            =   90
      TabIndex        =   5
      Top             =   45
      Width           =   6315
      Begin VB.PictureBox picBack 
         BorderStyle     =   0  'None
         Height          =   1635
         Left            =   45
         ScaleHeight     =   1635
         ScaleWidth      =   6180
         TabIndex        =   6
         Top             =   225
         Width           =   6180
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   2
            Left            =   2790
            TabIndex        =   3
            Text            =   "5"
            Top             =   1035
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   1
            Left            =   2790
            TabIndex        =   2
            Text            =   "1"
            Top             =   585
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Height          =   330
            Index           =   0
            Left            =   2790
            MaxLength       =   256
            TabIndex        =   1
            Top             =   135
            Width           =   3345
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Reconnection attempts:"
            Height          =   285
            Index           =   2
            Left            =   135
            TabIndex        =   9
            Top             =   1080
            Width           =   2625
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Minutes pause before Reconnect:"
            Height          =   285
            Index           =   1
            Left            =   135
            TabIndex        =   8
            Top             =   630
            Width           =   2625
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DUN Connection Name:"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   7
            Top             =   180
            Width           =   2670
         End
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   420
      Left            =   4860
      TabIndex        =   0
      Top             =   2070
      Width           =   1545
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   3240
      TabIndex        =   4
      Top             =   2070
      Width           =   1545
   End
End
Attribute VB_Name = "frmOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                           ByVal nIndex As Long, _
                                                                           ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                           ByVal nIndex As Long) As Long

Private Sub MakeFieldNumericOnly(ByVal lFieldHwnd As Long)

    Dim lRet As Long

    lRet = GetWindowLong(lFieldHwnd, -16)
    Call SetWindowLong(lFieldHwnd, -16, lRet Or &H2000)

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    Dim lRet As Long
    
    If LenB(Text1(0).Text) = 0 Then
        Beep
        MsgBox "You must supply the name of the Dial Up Networking connection (e.g. 'My Internet Connection') to be used.", vbExclamation
        Exit Sub
    End If
    
    strISP = Text1(0).Text
    lWaitBeforeReconnect = CLng(Text1(1).Text)
    lRetries = CLng(Text1(2).Text)
    
    lRet = WritePrivateProfileString("Options", "ISP", strISP, strINIFilePath)
    lRet = WritePrivateProfileString("Options", "Wait", CStr(lWaitBeforeReconnect), strINIFilePath)
    lRet = WritePrivateProfileString("Options", "Retries", CStr(lRetries), strINIFilePath)
    
    Unload Me

End Sub

Private Sub Form_Load()
        
    Call RestoreFormPosition(Me)
    
    Call MakeFieldNumericOnly(Text1(1).hwnd)
    Call MakeFieldNumericOnly(Text1(2).hwnd)
    Text1(0).Text = strISP
    Text1(1).Text = lWaitBeforeReconnect
    Text1(2).Text = lRetries

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Call SaveFormPosition(Me)

End Sub

Attribute VB_Name = "modReconn"
Option Explicit

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                                                 ByVal lpKeyName As Any, _
                                                                                                 ByVal lpDefault As String, _
                                                                                                 ByVal lpReturnedString As String, _
                                                                                                 ByVal nSize As Long, _
                                                                                                 ByVal lpFileName As String) As Long
                                                                                                 
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                                                     ByVal lpKeyName As Any, _
                                                                                                     ByVal lpString As Any, _
                                                                                                     ByVal lpFileName As String) As Long
Public strISP                   As String
Public lWaitBeforeReconnect     As Long
Public lRetries                 As Long
Public strINIFilePath           As String


Sub Main()
    
    Dim strINIRet As String * 256
    Dim lRet As Long
    
    
    strINIFilePath = App.Path & "\Reconnect.dat"
    lRet = GetPrivateProfileString("Options", "ISP", vbNullString, strINIRet, 256, strINIFilePath)
    strISP = Left$(strINIRet, lRet)
    lRet = GetPrivateProfileString("Options", "Wait", "0", strINIRet, 256, strINIFilePath)
    lWaitBeforeReconnect = CLng(Left$(strINIRet, lRet))
    lRet = GetPrivateProfileString("Options", "Retries", "5", strINIRet, 256, strINIFilePath)
    lRetries = CLng(Left$(strINIRet, lRet))
    
    frmMain.Show
    
End Sub

Public Sub SaveFormPosition(ByRef UForm As Form)

  Dim lRet As Long

    With UForm
        If .WindowState = vbMinimized Or .WindowState = vbMaximized Then
            Exit Sub
        End If
        If .Left > -1 Then
            If .Top > -1 Then
                If TypeOf UForm Is MDIForm Then
                    lRet = WritePrivateProfileString("FPos", .Tag, Format$(.Left, "00000") & Format$(.Top, "00000") & Format$(.Width, "00000") & Format$(.Height, "00000"), strINIFilePath)
                 ElseIf .BorderStyle = 2 Or .BorderStyle = 5 Then ' sizeable
                    lRet = WritePrivateProfileString("FPos", .Tag, Format$(.Left, "00000") & Format$(.Top, "00000") & Format$(.Width, "00000") & Format$(.Height, "00000"), strINIFilePath)
                 Else
                    lRet = WritePrivateProfileString("FPos", .Tag, Format$(.Left, "00000") & Format$(.Top, "00000"), strINIFilePath)
                End If
            End If
        End If
    End With

End Sub

Public Sub RestoreFormPosition(ByRef UForm As Form, Optional ByVal iNoMove As Integer = 0)

  Dim lRet      As Long
  Dim strINIRet As String * 25

    With UForm
        lRet = GetPrivateProfileString("FPos", .Tag, vbNullString, strINIRet, 25, strINIFilePath)
        If TypeOf UForm Is MDIForm Then
            If lRet = 20 Then
                .Move Mid$(strINIRet, 1, 5), Mid$(strINIRet, 6, 5), Mid$(strINIRet, 11, 5), Mid$(strINIRet, 16, 5)
             Else
                .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
            End If
         ElseIf .BorderStyle = 2 Or .BorderStyle = 5 Then ' sizeable
            If lRet = 20 Then
                .Move Mid$(strINIRet, 1, 5), Mid$(strINIRet, 6, 5), Mid$(strINIRet, 11, 5), Mid$(strINIRet, 16, 5)
             ElseIf iNoMove = 0 Then
                .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
            End If
         Else
            If lRet = 10 Then
                .Move Mid$(strINIRet, 1, 5), Mid$(strINIRet, 6, 5)
             ElseIf iNoMove = 0 Then
                .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
            End If
        End If
        If .Left > Screen.Width Or .Top > Screen.Height Then
            .Move (Screen.Width - .Width), (Screen.Height - .Height)
        End If
        If (.Left + .Width) < 0 Or (.Top + .Height) < 0 Then
            .Move 100, 100
        End If
    End With

End Sub



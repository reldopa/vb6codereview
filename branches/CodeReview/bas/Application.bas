Attribute VB_Name = "Application"
Option Explicit
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Public VBInstance As VBIDE.VBE
Public Connect As Connect
Public Const AppName = "CodeReview"

'-----------------------------------------------------------------------------
' Methods
'-----------------------------------------------------------------------------
Public Sub TrataErro(sRotina As String, sNomeModulo As String, Optional Display As Boolean)
    Dim strLog As String
     
    strLog = Date & ";" & Time & ";" & _
            sRotina & ";" & sNomeModulo & ";" & _
            strLog & Err.Description & ";" & _
            "At Line " & Erl
    
    Open "C:\errorLog.txt" For Append As #1
    Print #1, strLog
    Close #1
        
    If Display Then
        MsgBox strLog, vbCritical, AppName
    End If
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UIAddin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'-----------------------------------------------------------------------------
' General Declarations
'-----------------------------------------------------------------------------
Private strVBProjeto As String
Private strVBComponente As String
Private mobjForm As frmAddIn

'-----------------------------------------------------------------------------
' Properties
'-----------------------------------------------------------------------------
Public Property Get VBProjeto() As String
    VBProjeto = strVBProjeto
End Property
Public Property Let VBProjeto(ByVal vNewValue As String)
    strVBProjeto = vNewValue
End Property

Public Property Get VBComponente() As String
    VBComponente = strVBComponente
End Property
Public Property Let VBComponente(ByVal vNewValue As String)
    strVBComponente = vNewValue
End Property

'-----------------------------------------------------------------------------
' Methods
'-----------------------------------------------------------------------------
Private Sub Class_Initialize()
    Me.VBProjeto = VBInstance.ActiveVBProject.Name
    Me.VBComponente = VBInstance.ActiveVBProject.VBComponents.Item(1).Name
    
End Sub

Private Sub Class_Terminate()
    Unload mobjForm
End Sub

Public Sub Show()
On Error GoTo Catch
    If mobjForm Is Nothing Then
        Set mobjForm = New frmAddIn
    End If
  
    Dim i As Integer
    i = ""
    mobjForm.Show
            
Finally:
    Exit Sub

Catch:
    TrataErro "Show", "UIAddin"
    Resume Finally
End Sub

Sub Hide()
On Error GoTo Catch
    mobjForm.Hide
    Connect.Hide
    
Finally:
    Exit Sub

Catch:
    TrataErro "Hide", "UIAddin"
    Resume Finally
End Sub

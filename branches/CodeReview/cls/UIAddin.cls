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
On Error GoTo Catch
    Me.VBProjeto = VBInstance.ActiveVBProject.Name
    Me.VBComponente = VBInstance.ActiveVBProject.VBComponents.Item(1).Name
Finally:
    Exit Sub

Catch:
    TrataErro "Class_Initialize", "UIAddin"
    Resume Finally
End Sub

Private Sub Class_Terminate()
    If Not mobjForm Is Nothing Then
        Unload mobjForm
    End If
End Sub

Friend Sub Show()
On Error GoTo Catch
    If mobjForm Is Nothing Then
        Set mobjForm = New frmAddIn
    End If
  
    mobjForm.OBJUI = Me
    mobjForm.Show
            
Finally:
    Exit Sub

Catch:
    TrataErro "Show", "UIAddin", True
    Resume Finally
End Sub

Friend Sub Hide()
On Error GoTo Catch
    mobjForm.Hide
    Connect.Hide
    
Finally:
    Exit Sub

Catch:
    TrataErro "Hide", "UIAddin"
    Resume Finally
End Sub

Friend Function ContaLinhas() As Long
On Error GoTo Catch
    Dim objVBComponent As VBComponent
    'Define objVBComponent para o componente do programa informado
    Set objVBComponent = VBInstance.VBProjects.Item(Me.VBProjeto).VBComponents.Item(Me.VBComponente)
    'Atribui o numero de linhas usando o método - countoflines - do componente objVBComponent
    ContaLinhas = CLng(objVBComponent.CodeModule.CountOfLines)

Finally:
    Exit Function

Catch:
    TrataErro "ContaLinhas", "UIAddin"
    Resume Finally
End Function

Friend Sub VerificaPadroes(subjectString As String, strPattern As String)
On Error GoTo Catch
    Dim myRegExp As RegExp
    Dim myMatches As MatchCollection
    Dim myMatch As Match
    Set myRegExp = New RegExp
    myRegExp.IgnoreCase = True
    myRegExp.Global = True
    'myRegExp.Pattern = "Dim \w+ As \w+"
    myRegExp.Pattern = strPattern
    Set myMatches = myRegExp.Execute(subjectString)
    For Each myMatch In myMatches
      MsgBox (myMatch.Value)
    Next

Finally:
    Exit Sub

Catch:
    TrataErro "VerificaPadroes", "UIAddin"
    Resume Finally
End Sub

VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CodeReview"
   ClientHeight    =   4260
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRegExp 
      Caption         =   "RegExp"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtFind 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   3735
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblExpressao 
      Caption         =   "Expressão"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblNumeroDeLinhas 
      Caption         =   "Linhas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   1275
      Width           =   1455
   End
   Begin VB.Label lblNomeComponente 
      Caption         =   "Nome Componente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   795
      Width           =   2055
   End
   Begin VB.Label lblNomeProjeto 
      Caption         =   "Nome Projeto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   315
      Width           =   2055
   End
   Begin VB.Label lblNroDeLinhas 
      Caption         =   "Nro de Linhas"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1275
      Width           =   1095
   End
   Begin VB.Label lblComponente 
      Caption         =   "Componente"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   795
      Width           =   975
   End
   Begin VB.Label lblProjeto 
      Caption         =   "Projeto"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   315
      Width           =   615
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------
' General Declarations
'-----------------------------------------------------------------------------
Private mobjUI As UIAddin


'-----------------------------------------------------------------------------
' Properties
'-----------------------------------------------------------------------------
Public Property Get OBJUI() As UIAddin
    OBJUI = mobjUI
End Property

Public Property Let OBJUI(ByVal vNewValue As UIAddin)
    Set mobjUI = vNewValue
End Property


'-----------------------------------------------------------------------------
' Events
'-----------------------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo Catch
    Me.lblNomeProjeto.Caption = mobjUI.VBProjeto
    Me.lblNomeComponente.Caption = mobjUI.VBComponente
    Me.lblNumeroDeLinhas.Caption = Str(mobjUI.ContaLinhas)
    'Set The Form as TopMost
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, _
        SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    
Finally:
    Exit Sub

Catch:
    TrataErro "Form_Load", "frmAddIn"
    Resume Finally
End Sub

Private Sub cmdFind_Click()
    Dim objVBComponent As VBComponent
    Set objVBComponent = VBInstance.VBProjects.Item(mobjUI.VBProjeto).VBComponents.Item(mobjUI.VBComponente)
                
    If Len(txtFind.Text) > 0 Then
        If objVBComponent.CodeModule.Find(txtFind.Text, 1, 1, -1, -1, False, False, Me.chkRegExp.Value) Then
            MsgBox "Achou"
        End If
    End If
    
    objVBComponent.CodeModule.InsertLines 4, "'This is a comment"
                
End Sub

Private Sub CancelButton_Click()
    'esconde a janela do add-in
    mobjUI.Hide
End Sub

'Private Sub cmdContar_Click()
'
'    Me.lblNumeroDeLinhas.Caption = Str(mobjUI.ContaLinhas)
'
'    Dim objVBComponent As VBComponent
'    Dim lngNroLinhas As Long
'    Dim strModuleCode As String
'
'    'Define objVBComponent para o componente do programa informado
'    Set objVBComponent = VBInstance.VBProjects.Item(mobjUI.VBProjeto).VBComponents.Item(mobjUI.VBComponente)
'
'
'    'Atribui o numero de linhas usando o método - countoflines - do componente objVBComponent
'    lngNroLinhas = CLng(objVBComponent.CodeModule.CountOfLines)
'    lblNumeroDeLinhas.Caption = Str(lngNroLinhas)
'
'    strModuleCode = objVBComponent.CodeModule.Lines(1, lngNroLinhas)
'
'    'Grava em arquivo
'    Open "C:\arquivo.txt" For Output As #1
'    Print #1, strModuleCode
'    Close #1
'
'    VerificaPadroes strModuleCode
'
'End Sub



VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CodeReview"
   ClientHeight    =   10260
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   375
      Left            =   7080
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtFind 
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtNumeroDeLinhas 
      Height          =   405
      Left            =   6480
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdContar 
      Caption         =   "Contar"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   5640
      Width           =   1215
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   315
      Width           =   2055
   End
   Begin VB.Label lblNroDeLinhas 
      Caption         =   "Nro de Linhas"
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   1515
      Width           =   1095
   End
   Begin VB.Label lblComponente 
      Caption         =   "Componente"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   795
      Width           =   975
   End
   Begin VB.Label lblProjeto 
      Caption         =   "Projeto"
      Height          =   255
      Left            =   480
      TabIndex        =   3
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



'-----------------------------------------------------------------------------
' Events
'-----------------------------------------------------------------------------
Private Sub Form_Load()
    If mobjUI Is Nothing Then
        Set mobjUI = New UIAddin
    End If
    
    'Set The Form as TopMost
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, _
        SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    
    Me.lblNomeProjeto.Caption = mobjUI.VBProjeto
    Me.lblNomeComponente.Caption = mobjUI.VBComponente
    Me.txtNumeroDeLinhas.Text = vbNullString
End Sub

Private Sub cmdFind_Click()
    Dim objVBComponent As VBComponent
    Set objVBComponent = VBInstance.VBProjects.Item(mobjUI.VBProjeto).VBComponents.Item(mobjUI.VBComponente)
                
'    If objVBComponent.CodeModule.Find(txtFind.Text, 1, 1, -1, -1, False, False, False) Then
'        MsgBox "Achou"
'    End If
    
    objVBComponent.CodeModule.InsertLines 4, "'This is a comment"
                
End Sub



Private Sub CancelButton_Click()
    'esconde a janela do add-in
    mobjUI.Hide
End Sub

Private Sub cmdContar_Click()
    Dim objVBComponent As VBComponent
    Dim lngNroLinhas As Long
    Dim strModuleCode As String
    
    'Define objVBComponent para o componente do programa informado
    Set objVBComponent = VBInstance.VBProjects.Item(mobjUI.VBProjeto).VBComponents.Item(mobjUI.VBComponente)
                            
    
    'Atribui o numero de linhas usando o método - countoflines - do componente objVBComponent
    lngNroLinhas = CLng(objVBComponent.CodeModule.CountOfLines)
    txtNumeroDeLinhas.Text = Str(lngNroLinhas)
        
    strModuleCode = objVBComponent.CodeModule.Lines(1, lngNroLinhas)
    
    'Grava em arquivo
    Open "C:\arquivo.txt" For Output As #1
    Print #1, strModuleCode
    Close #1
    
    VerificaPadroes strModuleCode
    
End Sub

Sub VerificaPadroes(subjectString As String)
    Dim myRegExp As RegExp
    Dim myMatches As MatchCollection
    Dim myMatch As Match
    Set myRegExp = New RegExp
    myRegExp.IgnoreCase = True
    myRegExp.Global = True
    myRegExp.Pattern = "Dim \w+ As \w+"
    Set myMatches = myRegExp.Execute(subjectString)
    For Each myMatch In myMatches
      MsgBox (myMatch.Value)
    Next
End Sub

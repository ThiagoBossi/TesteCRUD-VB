VERSION 5.00
Begin VB.Form frmInicio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CRUD - Início"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSair 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   7455
   End
   Begin VB.CommandButton btnRealizarCadastro 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Realizar Cadastro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.CommandButton btnListarCadastros 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Listar Cadastros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnListarCadastros_Click()
    frmListarCadastros.Show 1
End Sub

Private Sub btnRealizarCadastro_Click()
    frmRealizarCadastro.Show 1
End Sub

Private Sub btnSair_Click()
    End
End Sub

Private Sub Form_Load()
    Call realizarConexao
    If Not frmLogin.fecharEntrar Then
        frmLogin.Show 1
    End If
End Sub

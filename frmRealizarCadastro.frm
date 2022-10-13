VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRealizarCadastro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CRUD - Realizar Cadastro"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Caption         =   "Preencha as Informações do Cadastro"
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10815
      Begin VB.CommandButton btnCancelar 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4440
         Width           =   4575
      End
      Begin VB.CommandButton btnSalvar 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Salvar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4440
         Width           =   4575
      End
      Begin MSComCtl2.DTPicker txtDataAniversario 
         Height          =   615
         Left            =   5520
         TabIndex        =   9
         Top             =   3600
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   1085
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   159383553
         CurrentDate     =   44839
      End
      Begin VB.ComboBox cbmOpcaoGenero 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   240
         TabIndex        =   8
         Top             =   3600
         Width           =   5055
      End
      Begin VB.TextBox txtCpfCadastro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5520
         TabIndex        =   6
         Top             =   2280
         Width           =   5055
      End
      Begin VB.TextBox txtTelefoneCadastro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   5055
      End
      Begin VB.TextBox txtNomeCadastro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   10335
      End
      Begin VB.Label Label 
         Caption         =   "Data de Aniversário:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   5520
         TabIndex        =   10
         Top             =   3120
         Width           =   5055
      End
      Begin VB.Label Label 
         Caption         =   "Opção de Gênero:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   3120
         Width           =   5055
      End
      Begin VB.Label Label 
         Caption         =   "CPF do Cadastro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5520
         TabIndex        =   5
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label Label 
         Caption         =   "Telefone do Cadastro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   3735
      End
      Begin VB.Label Label 
         Caption         =   "Nome do Cadastro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   10335
      End
   End
End
Attribute VB_Name = "frmRealizarCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnSalvar_Click()
    Dim SQL As String
    SQL = "INSERT INTO cadastros (nome, cpf, telefone, opcao, nascimento) VALUES ('" & txtNomeCadastro.Text & "', '" & txtCpfCadastro.Text & "', '" & txtTelefoneCadastro.Text & "', '" & cbmOpcaoGenero.ListIndex & "', '" & Format(txtDataAniversario.Value, "yyyy-MM-dd") & "')"
    cn.Execute SQL
    MsgBox "Cadastro realizado com sucesso!"
    Unload Me
    frmListarCadastros.Show 1
End Sub

Private Sub Form_Load()
    cbmOpcaoGenero.AddItem "Masculino"
    cbmOpcaoGenero.AddItem "Feminino"
    cbmOpcaoGenero.AddItem "Avião"
    
    frmLogin.fecharEntrar = True
End Sub

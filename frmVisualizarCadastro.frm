VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVisualizarCadastro 
   Caption         =   "CRUD - Visualizar Cadastro"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Caption         =   "Preencha as Informações do Cadastro"
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10815
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
         TabIndex        =   7
         Top             =   960
         Width           =   10335
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
         TabIndex        =   6
         Top             =   2280
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
         TabIndex        =   5
         Top             =   2280
         Width           =   5055
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
         TabIndex        =   4
         Top             =   3600
         Width           =   5055
      End
      Begin VB.CommandButton btnSalvar 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   2
         Top             =   4440
         Width           =   4575
      End
      Begin VB.CommandButton btnExcluir 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Excluir"
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
         TabIndex        =   1
         Top             =   4440
         Width           =   4575
      End
      Begin MSComCtl2.DTPicker txtDataAniversario 
         Height          =   615
         Left            =   5520
         TabIndex        =   3
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
         Format          =   158728193
         CurrentDate     =   44839
      End
      Begin VB.Label lblCodigo 
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
         Left            =   3840
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label 
         Caption         =   "Nome do Cadastro: #"
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
         Left            =   360
         TabIndex        =   12
         Top             =   480
         Width           =   3495
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
         TabIndex        =   11
         Top             =   1800
         Width           =   3735
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
         TabIndex        =   10
         Top             =   1800
         Width           =   3015
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
         TabIndex        =   9
         Top             =   3120
         Width           =   5055
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
         TabIndex        =   8
         Top             =   3120
         Width           =   5055
      End
   End
End
Attribute VB_Name = "frmVisualizarCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ssql As String
Dim rsCadastros As ADODB.Recordset

Private Sub btnExcluir_Click()
    If MsgBox("Você realmente deseja excluir este cadastro?", vbYesNo + vbDefaultButton2 + vbQuestion, "Exclusão de Cadastro") = vbYes Then
        ssql = "DELETE FROM cadastros WHERE codigo = '" & codigoCadastro & "'"
        cn.Execute ssql
        MsgBox "Cadastro excluído com sucesso..."
        Unload Me
    End If
    
End Sub

Private Sub btnSalvar_Click()
    ssql = "UPDATE cadastros SET nome = '" & txtNomeCadastro.Text & "', telefone = '" & txtTelefoneCadastro.Text & "', cpf = '" & txtCpfCadastro.Text & "', opcao = '" & cbmOpcaoGenero.ListIndex & "', nascimento = '" & Format(txtDataAniversario.Value, "yyyy-MM-dd") & "' WHERE codigo = '" & codigoCadastro & "'"
    cn.Execute ssql
    MsgBox "Cadastro atualizado com sucesso!"
End Sub

Private Sub Form_Load()
    lblCodigo.Caption = codigoCadastro
    ssql = "SELECT * FROM cadastros WHERE codigo = '" & codigoCadastro & "'"
    
    cbmOpcaoGenero.AddItem "Masculino"
    cbmOpcaoGenero.AddItem "Feminino"
    cbmOpcaoGenero.AddItem "Avião"
    
    Set rsCadastros = New ADODB.Recordset
    rsCadastros.Open ssql, cn, adOpenStatic
    
    If rsCadastros.RecordCount > 0 Then
        txtNomeCadastro.Text = rsCadastros!nome
        txtTelefoneCadastro.Text = rsCadastros!telefone
        txtCpfCadastro.Text = rsCadastros!cpf
        txtDataAniversario.Value = CDate(rsCadastros!nascimento)
        cbmOpcaoGenero.ListIndex = rsCadastros!opcao
    End If
End Sub


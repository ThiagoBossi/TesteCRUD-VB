VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmListarCadastros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CRUD - Listar Cadastros"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   12615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnVoltar 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Voltar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7320
      Width           =   6015
   End
   Begin VB.CommandButton btnNovoCadastro 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Realizar Novo Cadastro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
      Width           =   6015
   End
   Begin VB.Frame Frame 
      Caption         =   "Listagem de Cadastros"
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12375
      Begin MSFlexGridLib.MSFlexGrid listagemCadastros 
         Height          =   6015
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   10610
         _Version        =   393216
         Cols            =   7
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label 
         Caption         =   "Para visualizar um cadastro basta realizar um clique duplo em cima do mesmo."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   11775
      End
   End
End
Attribute VB_Name = "frmListarCadastros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnNovoCadastro_Click()
    frmRealizarCadastro.Show
    Unload Me
End Sub

Private Sub btnVoltar_Click()
    Unload Me
End Sub

Private Function exibirTabela()
    Dim ssql As String
    Dim rs As ADODB.Recordset
    
    ssql = "SELECT * FROM cadastros"
    
    Set rs = New ADODB.Recordset
    rs.Open ssql, cn, adOpenStatic
    
    listagemCadastros.Clear
    listagemCadastros.Rows = 1
    
    If rs.RecordCount <= 0 Then
        Exit Function
    End If
    
    rs.MoveFirst
    Do Until rs.EOF
        listagemCadastros.Rows = listagemCadastros.Rows + 1
        listagemCadastros.Row = listagemCadastros.Rows - 1
        
        If Not IsNull(rs!codigo) Then listagemCadastros.TextMatrix(listagemCadastros.Row, 1) = rs!codigo
        If Not IsNull(rs!nome) Then listagemCadastros.TextMatrix(listagemCadastros.Row, 2) = rs!nome
        If Not IsNull(rs!telefone) Then listagemCadastros.TextMatrix(listagemCadastros.Row, 3) = rs!telefone
        If Not IsNull(rs!cpf) Then listagemCadastros.TextMatrix(listagemCadastros.Row, 4) = rs!cpf
        If Not IsNull(rs!opcao) Then
            Select Case rs!opcao
                Case "0"
                    listagemCadastros.TextMatrix(listagemCadastros.Row, 5) = "Masculino"
                Case "1"
                    listagemCadastros.TextMatrix(listagemCadastros.Row, 5) = "Feminino"
                Case "2"
                    listagemCadastros.TextMatrix(listagemCadastros.Row, 5) = "Avião"
                Case Else
                    listagemCadastros.TextMatrix(listagemCadastros.Row, 5) = "N/A"
            End Select
        End If
        If Not IsNull(rs!nascimento) Then listagemCadastros.TextMatrix(listagemCadastros.Row, 6) = rs!nascimento
    
        listagemCadastros.FillStyle = flexFillRepeat
        listagemCadastros.Col = 1
        listagemCadastros.ColSel = listagemCadastros.Cols - 1
        listagemCadastros.FillStyle = flexFillSingle
        
        rs.MoveNext
    Loop
    
    listagemCadastros.ColWidth(0) = 800
    listagemCadastros.ColWidth(1) = 2000
    listagemCadastros.ColWidth(2) = 2000
    listagemCadastros.ColWidth(3) = 2000
    listagemCadastros.ColWidth(4) = 2000
    listagemCadastros.ColWidth(5) = 2000
    listagemCadastros.ColWidth(6) = 2000
    
    listagemCadastros.Row = 0
    listagemCadastros.Col = 1
    listagemCadastros.Text = "Código"
    listagemCadastros.Col = 2
    listagemCadastros.Text = "Nome"
    listagemCadastros.Col = 3
    listagemCadastros.Text = "Telefone"
    listagemCadastros.Col = 4
    listagemCadastros.Text = "CPF"
    listagemCadastros.Col = 5
    listagemCadastros.Text = "Gênero"
    listagemCadastros.Col = 6
    listagemCadastros.Text = "Data de Nascimento"
    
    rs.Close

    listagemCadastros.Redraw = True
End Function

Private Sub Form_Load()
    Call exibirTabela
End Sub

Private Sub listagemCadastros_DblClick()
    Dim codigo As String
    codigo = listagemCadastros.TextMatrix(listagemCadastros.Row, 1)
    codigoCadastro = codigo
    
    frmVisualizarCadastro.Show 1
    Call exibirTabela
End Sub

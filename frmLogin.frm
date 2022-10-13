VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CRUD - Login"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Caption         =   "Realize o Login:"
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4455
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
         Height          =   615
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3720
         Width           =   3495
      End
      Begin VB.CommandButton btnLogin 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Acessar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2760
         Width           =   3495
      End
      Begin VB.TextBox txtSenha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1920
         Width           =   3975
      End
      Begin VB.TextBox txtUsuario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "Senha:"
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
         TabIndex        =   4
         Top             =   1560
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Usuário:"
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
         Width           =   3975
      End
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Teste de CRUD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fecharEntrar As Boolean

Private Sub btnLogin_Click()
    Dim rsDados As ADODB.Recordset
    Dim SQL As String
    
    Set rsDados = New ADODB.Recordset
    SQL = "SELECT * FROM usuarios WHERE usuario = '" & txtUsuario.Text & "' AND senha = CONVERT(VARCHAR(32), HashBytes('MD5', '" & txtSenha.Text & "'))"
    rsDados.Open SQL, cn, adOpenStatic
    
    If rsDados.RecordCount > 0 Then
        fecharEntrar = True
        Unload Me
    Else
        MsgBox "Usuário ou senha incorretos..."
    End If
    
    Set rsDados = Nothing
End Sub

Private Sub btnSair_Click()
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not fecharEntrar Then
        End
    End If
End Sub

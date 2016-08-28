VERSION 5.00
Begin VB.Form frmFornecedores 
   Caption         =   "Cadastros"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   6720
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Localização"
      Height          =   810
      Left            =   255
      TabIndex        =   24
      Top             =   4245
      Width           =   5865
      Begin VB.TextBox txtLocalizar 
         Height          =   345
         Left            =   960
         TabIndex        =   8
         Top             =   285
         Width           =   4755
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   345
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdEditar 
      Appearance      =   0  'Flat
      Caption         =   "&Editar"
      Height          =   345
      Left            =   2700
      TabIndex        =   11
      Top             =   5655
      Width           =   1170
   End
   Begin VB.CommandButton cmdSalvar 
      Appearance      =   0  'Flat
      Caption         =   "&Salvar"
      Height          =   345
      Left            =   3900
      TabIndex        =   12
      Top             =   5655
      Width           =   1170
   End
   Begin VB.CommandButton cmdExcluir 
      Appearance      =   0  'Flat
      Caption         =   "E&xcluir"
      Height          =   345
      Left            =   1485
      TabIndex        =   10
      Top             =   5655
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5100
      TabIndex        =   13
      Top             =   5655
      Width           =   1170
   End
   Begin VB.CommandButton cmdAdicionar 
      Appearance      =   0  'Flat
      Caption         =   "&Adicionar"
      Height          =   345
      Left            =   270
      TabIndex        =   9
      Top             =   5655
      Width           =   1170
   End
   Begin VB.TextBox txtBairro 
      DataField       =   "Bairro"
      DataSource      =   "datFornecedores"
      Height          =   330
      Left            =   1365
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   3240
      Width           =   5040
   End
   Begin VB.TextBox txtCidade 
      DataField       =   "Cidade"
      DataSource      =   "datFornecedores"
      Height          =   330
      Left            =   1365
      TabIndex        =   6
      Top             =   3675
      Width           =   3585
   End
   Begin VB.TextBox txtUf 
      DataField       =   "UF"
      DataSource      =   "datFornecedores"
      Height          =   330
      Left            =   5985
      TabIndex        =   7
      Top             =   3675
      Width           =   420
   End
   Begin VB.TextBox txtEndereco 
      DataField       =   "Endereco"
      DataSource      =   "datFornecedores"
      Height          =   330
      Left            =   1335
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2820
      Width           =   5070
   End
   Begin VB.TextBox txtEmail 
      DataField       =   "EMail"
      DataSource      =   "datFornecedores"
      Height          =   330
      Left            =   1320
      TabIndex        =   3
      Top             =   2400
      Width           =   5070
   End
   Begin VB.TextBox txtTelefone 
      DataField       =   "Telefone"
      DataSource      =   "datFornecedores"
      Height          =   330
      Left            =   4770
      TabIndex        =   2
      Top             =   1965
      Width           =   1620
   End
   Begin VB.TextBox txtCgc 
      DataField       =   "CGC"
      DataSource      =   "datFornecedores"
      Height          =   330
      Left            =   1350
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1980
      Width           =   2355
   End
   Begin VB.TextBox txtRazaoSocial 
      DataField       =   "RazaoSocial"
      DataSource      =   "datFornecedores"
      Height          =   330
      Left            =   1350
      TabIndex        =   0
      Top             =   1560
      Width           =   5025
   End
   Begin VB.TextBox txtCodigo 
      DataField       =   "Codigo"
      DataSource      =   "datFornecedores"
      Height          =   330
      Left            =   1350
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1395
   End
   Begin VB.Data datFornecedores 
      Appearance      =   0  'Flat
      Caption         =   "Navega entre os registros."
      Connect         =   "Access"
      DatabaseName    =   "C:\Alunos\Tarde\Estoque\estoque.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2910
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblFornecedores"
      Top             =   5175
      Width           =   3225
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FORNECEDORES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   480
      TabIndex        =   26
      Top             =   300
      Width           =   3240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Endereço:"
      Height          =   195
      Left            =   285
      TabIndex        =   23
      Top             =   2835
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail:"
      Height          =   195
      Left            =   525
      TabIndex        =   22
      Top             =   2460
      Width           =   480
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Cidade:"
      Height          =   195
      Left            =   510
      TabIndex        =   18
      Top             =   3720
      Width           =   540
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "UF.:"
      Height          =   195
      Left            =   5370
      TabIndex        =   19
      Top             =   3750
      Width           =   300
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Bairro:"
      Height          =   195
      Left            =   600
      TabIndex        =   21
      Top             =   3285
      Width           =   450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Telefone:"
      Height          =   195
      Left            =   3855
      TabIndex        =   20
      Top             =   2025
      Width           =   675
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   495
      TabIndex        =   17
      Top             =   1155
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "C.G.C.:"
      Height          =   195
      Left            =   525
      TabIndex        =   16
      Top             =   2025
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "R. Social:"
      Height          =   195
      Left            =   345
      TabIndex        =   15
      Top             =   1605
      Width           =   690
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FORNECEDORES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   495
      TabIndex        =   27
      Top             =   330
      Width           =   3240
   End
End
Attribute VB_Name = "frmFornecedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdAdicionar_Click()
    Trava False
    Me.datFornecedores.Recordset.AddNew
    Me.txtRazaoSocial.SetFocus
End Sub

Private Sub cmdCancelar_Click()
    cmdEditar_Click
    If Me.datFornecedores.EditMode = dbEditNone Then Exit Sub
    Me.datFornecedores.Recordset.CancelUpdate
    Trava True
End Sub

Private Sub cmdEditar_Click()
    If Me.datFornecedores.Recordset.RecordCount = 0 Then Exit Sub
    Trava False
    Me.datFornecedores.Recordset.Edit
End Sub

Private Sub cmdExcluir_Click()
    Dim MSG As String
    MSG = "Deseja excluir ?"
    If Me.datFornecedores.Recordset.RecordCount = 0 Then Exit Sub
    If MsgBox(MSG, vbQuestion + vbYesNo, "Excluir") = vbNo Then Exit Sub
    Me.datFornecedores.Recordset.Delete
    If Me.datFornecedores.Recordset.RecordCount = 0 Then
        Me.datFornecedores.Refresh
    Else
        Me.datFornecedores.Recordset.MoveNext
        If Me.datFornecedores.Recordset.EOF Then
            Me.datFornecedores.Recordset.MoveLast
        End If
    End If
End Sub

Private Sub cmdSalvar_Click()
    If Me.datFornecedores.EditMode = dbEditNone Then Exit Sub
    Me.datFornecedores.Recordset.Update
    Me.datFornecedores.Recordset.Bookmark = _
        Me.datFornecedores.Recordset.LastModified
    Trava True
End Sub

Private Sub datFornecedores_Reposition()
    On Error Resume Next
    Me.datFornecedores.Caption = "Registro " & Me.datFornecedores.Recordset.AbsolutePosition + 1
End Sub

Private Sub Form_Activate()
        Trava True
End Sub

Private Sub Form_Load()
    CaminhoBanco Me
End Sub



Private Sub txtLocalizar_Change()
    If Me.datFornecedores.Recordset.RecordCount = 0 Then Exit Sub
   Me.datFornecedores.Recordset.FindFirst "RazaoSocial like '*" & Me.txtLocalizar.Text & "*'"
End Sub


Private Sub Trava(Sim As Boolean)
    Me.txtRazaoSocial.Locked = Sim
    Me.txtCgc.Locked = Sim
    Me.txtTelefone.Locked = Sim
    Me.txtEmail.Locked = Sim
    Me.txtEndereco.Locked = Sim
    Me.txtBairro.Locked = Sim
    Me.txtCidade.Locked = Sim
    Me.txtUf.Locked = Sim
    Me.txtLocalizar.Locked = Not Sim
End Sub




VERSION 5.00
Begin VB.Form frmClientes 
   Caption         =   "Cadastros"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   7665
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "VOLTAR"
      Height          =   255
      Left            =   5010
      TabIndex        =   32
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MARCAR"
      Height          =   300
      Left            =   3225
      TabIndex        =   31
      Top             =   255
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      Caption         =   "Localização"
      Height          =   810
      Left            =   405
      TabIndex        =   27
      Top             =   4290
      Width           =   6855
      Begin VB.TextBox txtLocalizar 
         Height          =   345
         Left            =   960
         TabIndex        =   9
         Top             =   300
         Width           =   5715
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   345
         Width           =   765
      End
   End
   Begin VB.ComboBox cmdSexo 
      DataField       =   "Sexo"
      DataSource      =   "datClientes"
      Height          =   315
      ItemData        =   "frmClientes.frx":0000
      Left            =   5100
      List            =   "frmClientes.frx":000A
      TabIndex        =   6
      Top             =   3015
      Width           =   1890
   End
   Begin VB.ComboBox cmbEstadoCivil 
      DataField       =   "Estado Civil"
      DataSource      =   "datClientes"
      Height          =   315
      ItemData        =   "frmClientes.frx":0023
      Left            =   1530
      List            =   "frmClientes.frx":002D
      TabIndex        =   5
      Top             =   3015
      Width           =   1890
   End
   Begin VB.CommandButton cmdAdicionar 
      Appearance      =   0  'Flat
      Caption         =   "&Adicionar"
      Height          =   345
      Left            =   750
      TabIndex        =   10
      Top             =   5685
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5580
      TabIndex        =   14
      Top             =   5685
      Width           =   1170
   End
   Begin VB.CommandButton cmdExcluir 
      Appearance      =   0  'Flat
      Caption         =   "E&xcluir"
      Height          =   345
      Left            =   1965
      TabIndex        =   11
      Top             =   5685
      Width           =   1170
   End
   Begin VB.CommandButton cmdSalvar 
      Appearance      =   0  'Flat
      Caption         =   "&Salvar"
      Height          =   345
      Left            =   4380
      TabIndex        =   13
      Top             =   5685
      Width           =   1170
   End
   Begin VB.CommandButton cmdEditar 
      Appearance      =   0  'Flat
      Caption         =   "&Editar"
      Height          =   345
      Left            =   3180
      TabIndex        =   12
      Top             =   5685
      Width           =   1170
   End
   Begin VB.TextBox txtNascimento 
      DataField       =   "Data de Nascimento"
      DataSource      =   "datClientes"
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   3405
      Width           =   1650
   End
   Begin VB.TextBox txtComentario 
      DataField       =   "Comentario"
      DataSource      =   "datClientes"
      Height          =   315
      Left            =   1560
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3810
      Width           =   5415
   End
   Begin VB.TextBox txtEndereco 
      DataField       =   "Endereco"
      DataSource      =   "datClientes"
      Height          =   315
      Left            =   1530
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2610
      Width           =   5460
   End
   Begin VB.TextBox txtTelefone 
      DataField       =   "Telefone"
      DataSource      =   "datClientes"
      Height          =   315
      Left            =   4770
      TabIndex        =   3
      Top             =   2205
      Width           =   2220
   End
   Begin VB.TextBox txtRG 
      DataField       =   "RG"
      DataSource      =   "datClientes"
      Height          =   315
      Left            =   1530
      TabIndex        =   2
      Top             =   2190
      Width           =   2220
   End
   Begin VB.Data datClientes 
      Appearance      =   0  'Flat
      Caption         =   "Navega entre os registros."
      Connect         =   "Access"
      DatabaseName    =   "D:\Dados\PROG\VBASIC\Estoque\estoque.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4050
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblClientes"
      Top             =   5205
      Width           =   3225
   End
   Begin VB.TextBox txtCodigo 
      DataField       =   "Codigo"
      DataSource      =   "datClientes"
      Height          =   315
      Left            =   1545
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1050
      Width           =   1215
   End
   Begin VB.TextBox txtNome 
      DataField       =   "Nome"
      DataSource      =   "datClientes"
      Height          =   315
      Left            =   1530
      TabIndex        =   0
      Top             =   1425
      Width           =   5460
   End
   Begin VB.TextBox txtSobrenome 
      DataField       =   "Sobrenome"
      DataSource      =   "datClientes"
      Height          =   315
      Left            =   1530
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1800
      Width           =   5460
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTES"
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
      Left            =   675
      TabIndex        =   30
      Top             =   285
      Width           =   1905
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTES"
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
      Left            =   690
      TabIndex        =   29
      Top             =   315
      Width           =   1905
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Comentário:"
      Height          =   195
      Left            =   540
      TabIndex        =   26
      Top             =   3885
      Width           =   840
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Dt. Nasc.:"
      Height          =   195
      Left            =   630
      TabIndex        =   25
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Sexo:"
      Height          =   195
      Left            =   4485
      TabIndex        =   24
      Top             =   3075
      Width           =   405
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Est. Cilvil.:"
      Height          =   195
      Left            =   645
      TabIndex        =   23
      Top             =   3075
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "R.G.:"
      Height          =   195
      Left            =   990
      TabIndex        =   22
      Top             =   2235
      Width           =   375
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Endereço:"
      Height          =   195
      Left            =   630
      TabIndex        =   21
      Top             =   2655
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Telefone:"
      Height          =   195
      Left            =   3930
      TabIndex        =   20
      Top             =   2250
      Width           =   675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   825
      TabIndex        =   19
      Top             =   1095
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CÓDIGO.:"
      Height          =   0
      Left            =   15
      TabIndex        =   18
      Top             =   315
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
      Height          =   195
      Left            =   900
      TabIndex        =   17
      Top             =   1470
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Sobrenome:"
      Height          =   195
      Left            =   510
      TabIndex        =   16
      Top             =   1815
      Width           =   855
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdicionar_Click()
    Trava False
    Me.datClientes.Recordset.AddNew
    Me.txtNome.SetFocus
End Sub

Private Sub cmdCancelar_Click()
    cmdEditar_Click
    If Me.datClientes.EditMode = dbEditNone Then Exit Sub
    Me.datClientes.Recordset.CancelUpdate
    Trava True
End Sub

Private Sub cmdEditar_Click()
    If Me.datClientes.Recordset.RecordCount = 0 Then Exit Sub
    Trava False
    Me.datClientes.Recordset.Edit
End Sub

Private Sub cmdExcluir_Click()
    Dim MSG As String
    MSG = "Deseja excluir ?"
    If Me.datClientes.Recordset.RecordCount = 0 Then Exit Sub
    If MsgBox(MSG, vbQuestion + vbYesNo, "Excluir") = vbNo Then Exit Sub
    Me.datClientes.Recordset.Delete
    If Me.datClientes.Recordset.RecordCount = 0 Then
        Me.datClientes.Refresh
    Else
        Me.datClientes.Recordset.MoveNext
        If Me.datClientes.Recordset.EOF Then
            Me.datClientes.Recordset.MoveLast
        End If
    End If
End Sub

Private Sub cmdSalvar_Click()
    If Me.datClientes.EditMode = dbEditNone Then Exit Sub
    Me.datClientes.Recordset.Update
    Me.datClientes.Recordset.Bookmark = _
        Me.datClientes.Recordset.LastModified
    Trava True
End Sub

Private Sub Command1_Click()
marca = Me.datClientes.Recordset.Bookmark
End Sub

Private Sub Command2_Click()
Me.datClientes.Recordset.Bookmark = marca
End Sub

Private Sub datClientes_Reposition()
    On Error Resume Next
    Me.datClientes.Caption = "Registro " & Me.datClientes.Recordset.AbsolutePosition + 1
End Sub

Private Sub Form_Activate()
        Trava True
End Sub

Private Sub Form_Load()
    CaminhoBanco Me
End Sub



Private Sub txtLocalizar_Change()
    If Me.datClientes.Recordset.RecordCount = 0 Then Exit Sub
   Me.datClientes.Recordset.FindFirst "Nome like '*" & Me.txtLocalizar.Text & "*'"
End Sub
Private Sub Trava(Sim As Boolean)
    Me.txtNome.Locked = Sim
    Me.txtSobrenome.Locked = Sim
    Me.txtRG.Locked = Sim
    Me.txtTelefone.Locked = Sim
    Me.txtEndereco.Locked = Sim
    Me.cmbEstadoCivil.Locked = Sim
    Me.cmdSexo.Locked = Sim
    Me.txtNascimento.Locked = Sim
    Me.txtComentario.Locked = Sim
    Me.txtLocalizar.Locked = Not Sim
End Sub




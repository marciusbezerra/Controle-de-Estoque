VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmProdutos 
   Caption         =   "Produtos."
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   7125
   WindowState     =   2  'Maximized
   Begin MSMask.MaskEdBox txtPreçoUnitário 
      DataField       =   "PrecoUnitario"
      DataSource      =   "datProdutos"
      Height          =   315
      Left            =   1635
      TabIndex        =   5
      Top             =   2880
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   556
      _Version        =   327681
      Format          =   "R$ #,##0.00;(R$ #,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Localização"
      Height          =   810
      Left            =   525
      TabIndex        =   20
      Top             =   3330
      Width           =   5865
      Begin VB.TextBox txtLocalizar 
         Height          =   330
         Left            =   960
         TabIndex        =   6
         Top             =   300
         Width           =   4755
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   345
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdAdicionar 
      Appearance      =   0  'Flat
      Caption         =   "&Adicionar"
      Height          =   345
      Left            =   450
      TabIndex        =   7
      Top             =   4785
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5280
      TabIndex        =   11
      Top             =   4785
      Width           =   1170
   End
   Begin VB.CommandButton cmdExcluir 
      Appearance      =   0  'Flat
      Caption         =   "E&xcluir"
      Height          =   345
      Left            =   1665
      TabIndex        =   8
      Top             =   4785
      Width           =   1170
   End
   Begin VB.CommandButton cmdSalvar 
      Appearance      =   0  'Flat
      Caption         =   "&Salvar"
      Height          =   345
      Left            =   4080
      TabIndex        =   10
      Top             =   4785
      Width           =   1170
   End
   Begin VB.CommandButton cmdEditar 
      Appearance      =   0  'Flat
      Caption         =   "&Editar"
      Height          =   345
      Left            =   2880
      TabIndex        =   9
      Top             =   4785
      Width           =   1170
   End
   Begin VB.ComboBox cmdUnidade 
      Height          =   315
      ItemData        =   "frmProdutos.frx":0000
      Left            =   5580
      List            =   "frmProdutos.frx":000D
      TabIndex        =   2
      Top             =   2010
      Width           =   1140
   End
   Begin MSDBCtls.DBCombo dbcCategoria 
      Bindings        =   "frmProdutos.frx":001D
      DataField       =   "Categoria"
      DataSource      =   "datProdutos"
      Height          =   315
      Left            =   1650
      TabIndex        =   1
      Top             =   1995
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   327681
      Style           =   2
      ListField       =   "Descricao"
      BoundColumn     =   "Codigo"
      Text            =   ""
   End
   Begin VB.Data datCombo 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Dados\PROG\VBASIC\Estoque\estoque.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3060
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblCategorias"
      Top             =   2025
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.TextBox txtEstoqueMinimo 
      DataField       =   "EstoqueMinimo"
      DataSource      =   "datProdutos"
      Height          =   315
      Left            =   5010
      TabIndex        =   4
      Top             =   2445
      Width           =   1695
   End
   Begin VB.Data datProdutos 
      Appearance      =   0  'Flat
      Caption         =   "Navega entre os registros."
      Connect         =   "Access"
      DatabaseName    =   "D:\Dados\PROG\VBASIC\Estoque\estoque.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3165
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblProdutos"
      Top             =   4245
      Width           =   3225
   End
   Begin VB.TextBox txtCodigo 
      DataField       =   "Codigo"
      DataSource      =   "datProdutos"
      Height          =   315
      Left            =   1635
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1155
      Width           =   1470
   End
   Begin VB.TextBox txtDescricao 
      DataField       =   "Descricao"
      DataSource      =   "datProdutos"
      Height          =   315
      Left            =   1635
      TabIndex        =   0
      Top             =   1560
      Width           =   5070
   End
   Begin VB.TextBox txtQuantidadedeEstoque 
      DataField       =   "QuantidadeEstoque"
      DataSource      =   "datProdutos"
      Height          =   315
      Left            =   1650
      TabIndex        =   3
      Top             =   2430
      Width           =   1695
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUTOS"
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
      Left            =   765
      TabIndex        =   22
      Top             =   315
      Width           =   2190
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Prç. Unitário:"
      Height          =   195
      Left            =   480
      TabIndex        =   19
      Top             =   2910
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Estoque Mínimo:"
      Height          =   195
      Left            =   3510
      TabIndex        =   18
      Top             =   2505
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
      Height          =   195
      Left            =   630
      TabIndex        =   17
      Top             =   1635
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   855
      TabIndex        =   16
      Top             =   1185
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Unidade:"
      Height          =   195
      Left            =   4710
      TabIndex        =   15
      Top             =   2070
      Width           =   645
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Qtd. Estoque:"
      Height          =   195
      Left            =   315
      TabIndex        =   14
      Top             =   2490
      Width           =   975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Categoria:"
      Height          =   195
      Left            =   660
      TabIndex        =   13
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUTOS"
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
      Left            =   780
      TabIndex        =   23
      Top             =   345
      Width           =   2190
   End
End
Attribute VB_Name = "frmProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdicionar_Click()
    Trava False
    Me.datProdutos.Recordset.AddNew
    Me.txtDescricao.SetFocus
End Sub

Private Sub cmdCancelar_Click()
    cmdEditar_Click
    If Me.datProdutos.EditMode = dbEditNone Then Exit Sub
    Me.datProdutos.Recordset.CancelUpdate
    Trava True
End Sub

Private Sub cmdEditar_Click()
    If Me.datProdutos.Recordset.RecordCount = 0 Then Exit Sub
    Trava False
    Me.datProdutos.Recordset.Edit
End Sub

Private Sub cmdExcluir_Click()
    Dim MSG As String
    MSG = "Deseja excluir ?"
    If Me.datProdutos.Recordset.RecordCount = 0 Then Exit Sub
    If MsgBox(MSG, vbQuestion + vbYesNo, "Excluir") = vbNo Then Exit Sub
    Me.datProdutos.Recordset.Delete
    If Me.datProdutos.Recordset.RecordCount = 0 Then
        Me.datProdutos.Refresh
    Else
        Me.datProdutos.Recordset.MoveNext
        If Me.datProdutos.Recordset.EOF Then
            Me.datProdutos.Recordset.MoveLast
        End If
    End If
End Sub

Private Sub cmdSalvar_Click()
    If Me.datProdutos.EditMode = dbEditNone Then Exit Sub
    Me.datProdutos.Recordset.Update
    Me.datProdutos.Recordset.Bookmark = _
        Me.datProdutos.Recordset.LastModified
    Trava True
End Sub

Private Sub datProdutos_Reposition()
    On Error Resume Next
    Me.datProdutos.Caption = "Registro " & Me.datProdutos.Recordset.AbsolutePosition + 1
End Sub


Private Sub Form_Activate()
        Trava True
End Sub

Private Sub Form_Load()
    CaminhoBanco Me
End Sub



Private Sub txtLocalizar_Change()
    If Me.datProdutos.Recordset.RecordCount = 0 Then Exit Sub
   Me.datProdutos.Recordset.FindFirst "Descricao like '*" & Me.txtLocalizar.Text & "*'"
End Sub


Private Sub Trava(Sim As Boolean)
    Me.txtDescricao.Locked = Sim
    Me.dbcCategoria.Locked = Sim
    Me.cmdUnidade.Locked = Sim
    Me.txtQuantidadedeEstoque.Locked = Sim
    Me.txtEstoqueMinimo.Locked = Sim
    Me.txtPreçoUnitário.Enabled = Not Sim
    Me.txtLocalizar.Locked = Not Sim
End Sub




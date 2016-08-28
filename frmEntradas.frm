VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmEntradas 
   Caption         =   "Histórico de entradas."
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   7245
   WindowState     =   2  'Maximized
   Begin VB.Data datCombo 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Alunos\tarde\ESTOQUE\estoque.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   465
      Left            =   3300
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblHistoricoDeEntradas"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Text4 
      DataField       =   "ValorDaVenda"
      DataSource      =   "datCorrige"
      Height          =   390
      Left            =   975
      TabIndex        =   10
      Top             =   2625
      Width           =   1515
   End
   Begin VB.TextBox Text3 
      DataField       =   "DataDaVenda"
      DataSource      =   "datCorrige"
      Height          =   390
      Left            =   975
      TabIndex        =   8
      Top             =   2175
      Width           =   1290
   End
   Begin VB.TextBox Text2 
      DataField       =   "Quantidade"
      DataSource      =   "datCorrige"
      Height          =   390
      Left            =   975
      TabIndex        =   6
      Top             =   1650
      Width           =   2715
   End
   Begin VB.TextBox Text1 
      DataField       =   "Funcionario"
      DataSource      =   "datCorrige"
      Height          =   390
      Left            =   975
      TabIndex        =   4
      Top             =   1200
      Width           =   3990
   End
   Begin VB.Data datCorrige 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Alunos\tarde\ESTOQUE\estoque.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   150
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblHistoricoDeEntradas"
      Top             =   3225
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox txtCodigo 
      DataField       =   "Codigo"
      DataSource      =   "datCorrige"
      Height          =   330
      Left            =   975
      TabIndex        =   1
      Top             =   225
      Width           =   1440
   End
   Begin MSDBCtls.DBCombo txtProduto 
      Bindings        =   "frmEntradas.frx":0000
      DataField       =   "Produto"
      DataSource      =   "datCorrige"
      Height          =   315
      Left            =   975
      TabIndex        =   11
      Top             =   750
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   556
      _Version        =   327681
      ListField       =   "Produto"
      BoundColumn     =   "Codigo"
      Text            =   ""
   End
   Begin VB.Label Label6 
      Caption         =   "Valor Venda.:"
      Height          =   240
      Left            =   0
      TabIndex        =   9
      Top             =   2775
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Data Venda.:"
      Height          =   195
      Left            =   45
      TabIndex        =   7
      Top             =   2325
      Width           =   945
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade.:"
      Height          =   195
      Left            =   75
      TabIndex        =   5
      Top             =   1800
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Funcionario.:"
      Height          =   195
      Left            =   75
      TabIndex        =   3
      Top             =   1350
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Produto.:"
      Height          =   240
      Left            =   225
      TabIndex        =   2
      Top             =   825
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Codigo.:"
      Height          =   195
      Left            =   405
      TabIndex        =   0
      Top             =   375
      Width           =   585
   End
End
Attribute VB_Name = "frmEntradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdicionar_Click()
    Me.txtProduto.SetFocus
    Me.datCorrige.Recordset.AddNew
    Exit Sub
End Sub

Private Sub cmdCancelar_Click()
    On Error Resume Next
    Me.datCorrige.Recordset.Edit
    Me.datCorrige.Recordset.CancelUpdate
End Sub

Private Sub cmdExcluir_Click()
 If Not Me.datCorrige.Recordset.BOF And Not Me.datCorrige.Recordset.EOF Then
       Me.datCorrige.Recordset.Delete
       Me.datCorrige.Recordset.MovePrevious
    If Me.datCorrige.Recordset.BOF Then Me.datCorrige.Recordset.MoveNext
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'   MsgBox KeyAscii
   If KeyAscii = 13 And TypeOf Screen.ActiveControl Is TextBox Then
      KeyAscii = 0
      SendKeys "{TAB}"
   End If
   If KeyAscii = 43 And TypeOf Screen.ActiveControl Is TextBox Then
      KeyAscii = 0
      SendKeys "+{TAB}"
   End If
End Sub

Private Sub Form_Load()
    CaminhoBanco Me
End Sub

VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSaidas 
   Caption         =   "Saidas de Produtos"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab tabVendas 
      Height          =   5820
      Left            =   105
      TabIndex        =   0
      Top             =   750
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   10266
      _Version        =   327681
      TabOrientation  =   3
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   706
      TabMaxWidth     =   35
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Dados"
      TabPicture(0)   =   "frmSaidas.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dbcFuncionario"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dbcClientes"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "datVendas"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdEditar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdSalvar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdExcluir"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdCancelar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdAdicionar"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtDTVenda"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "datListClientes"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "datListfuncionarios"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Vendas"
      TabPicture(1)   =   "frmSaidas.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label9"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdDeletar"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdAddProd"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdVisualizar"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtPreco"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtQtd"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      Begin VB.TextBox txtQtd 
         Height          =   330
         Left            =   3270
         TabIndex        =   27
         Top             =   2850
         Width           =   600
      End
      Begin MSMask.MaskEdBox txtPreco 
         Height          =   330
         Left            =   1125
         TabIndex        =   25
         Top             =   2850
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   327681
         Format          =   "R$ #,##0.00;(R$ #,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdVisualizar 
         Caption         =   "&Visualizar"
         Height          =   330
         Left            =   6525
         TabIndex        =   24
         Top             =   2850
         Width           =   1290
      End
      Begin VB.CommandButton cmdAddProd 
         Caption         =   "&Adicionar"
         Height          =   330
         Left            =   4485
         TabIndex        =   23
         Top             =   2850
         Width           =   975
      End
      Begin VB.CommandButton cmdDeletar 
         Caption         =   "&Limpar"
         Height          =   330
         Left            =   5505
         TabIndex        =   22
         Top             =   2850
         Width           =   750
      End
      Begin VB.Frame Frame3 
         Caption         =   "Produtos faturados"
         Height          =   2025
         Left            =   195
         TabIndex        =   18
         Top             =   3240
         Width           =   7680
         Begin VB.Data datFaturados 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   "D:\Dados\PROG\VBASIC\Estoque\estoque.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   435
            Left            =   6165
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "tblHistoricoDeVenda"
            Top             =   1395
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSDBGrid.DBGrid grdFaturados 
            Bindings        =   "frmSaidas.frx":0038
            Height          =   1620
            Left            =   165
            OleObjectBlob   =   "frmSaidas.frx":004F
            TabIndex        =   28
            Top             =   285
            Width           =   7350
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Localizar produtos"
         Height          =   690
         Left            =   135
         TabIndex        =   14
         Top             =   195
         Width           =   7590
         Begin VB.TextBox txtProduto 
            Height          =   285
            Left            =   1170
            TabIndex        =   15
            Top             =   270
            Width           =   6300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   180
            TabIndex        =   16
            Top             =   285
            Width           =   765
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Produtos em estoque"
         Height          =   1785
         Left            =   180
         TabIndex        =   12
         Top             =   945
         Width           =   7680
         Begin VB.Data DatProduto 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   "D:\Dados\PROG\VBASIC\Estoque\estoque.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   435
            Left            =   6165
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "ProdutosEstocados"
            Top             =   1140
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSDBGrid.DBGrid grdProdutos 
            Bindings        =   "frmSaidas.frx":0BEF
            Height          =   1410
            Left            =   150
            OleObjectBlob   =   "frmSaidas.frx":0C04
            TabIndex        =   13
            Top             =   270
            Width           =   7350
         End
      End
      Begin VB.Data datListfuncionarios 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "D:\Dados\PROG\VBASIC\Estoque\estoque.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   -69000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "cboFuncionarios"
         Top             =   1290
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data datListClientes 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "D:\Dados\PROG\VBASIC\Estoque\estoque.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   -69015
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "cboClientes"
         Top             =   1665
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtDTVenda 
         DataField       =   "DataDaVenda"
         DataSource      =   "datVendas"
         Height          =   315
         Left            =   -73305
         TabIndex        =   6
         Top             =   2025
         Width           =   1515
      End
      Begin VB.CommandButton cmdAdicionar 
         Appearance      =   0  'Flat
         Caption         =   "&Adicionar"
         Height          =   345
         Left            =   -74280
         TabIndex        =   5
         Top             =   3195
         Width           =   1170
      End
      Begin VB.CommandButton cmdCancelar 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   -69450
         TabIndex        =   4
         Top             =   3195
         Width           =   1170
      End
      Begin VB.CommandButton cmdExcluir 
         Appearance      =   0  'Flat
         Caption         =   "E&xcluir"
         Height          =   345
         Left            =   -73065
         TabIndex        =   3
         Top             =   3195
         Width           =   1170
      End
      Begin VB.CommandButton cmdSalvar 
         Appearance      =   0  'Flat
         Caption         =   "&Salvar"
         Height          =   345
         Left            =   -70650
         TabIndex        =   2
         Top             =   3195
         Width           =   1170
      End
      Begin VB.CommandButton cmdEditar 
         Appearance      =   0  'Flat
         Caption         =   "&Editar"
         Height          =   345
         Left            =   -71850
         TabIndex        =   1
         Top             =   3195
         Width           =   1170
      End
      Begin VB.Data datVendas 
         Appearance      =   0  'Flat
         Caption         =   "Navega entre os registros."
         Connect         =   "Access"
         DatabaseName    =   "D:\Dados\PROG\VBASIC\Estoque\estoque.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -71565
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "tblVendas"
         Top             =   2655
         Width           =   3225
      End
      Begin MSDBCtls.DBCombo dbcClientes 
         Bindings        =   "frmSaidas.frx":1963
         DataField       =   "CodCliente"
         DataSource      =   "datVendas"
         Height          =   315
         Left            =   -73320
         TabIndex        =   7
         Top             =   1665
         Width           =   4350
         _ExtentX        =   7673
         _ExtentY        =   556
         _Version        =   327681
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "NomeCompleto"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSDBCtls.DBCombo dbcFuncionario 
         Bindings        =   "frmSaidas.frx":197D
         DataField       =   "CodFuncionario"
         DataSource      =   "datVendas"
         Height          =   315
         Left            =   -73305
         TabIndex        =   8
         Top             =   1305
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   327681
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Apelido"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Qtd.:"
         Height          =   195
         Left            =   2775
         TabIndex        =   26
         Top             =   2895
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Prç. Venda:"
         Height          =   195
         Left            =   210
         TabIndex        =   21
         Top             =   2895
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Preço  de Venda : "
         Height          =   195
         Left            =   135
         TabIndex        =   17
         Top             =   1410
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Funcionário :"
         Height          =   195
         Left            =   -74385
         TabIndex        =   11
         Top             =   1365
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Clientes :"
         Height          =   195
         Left            =   -74115
         TabIndex        =   10
         Top             =   1710
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Dt. Venda:"
         Height          =   195
         Left            =   -74235
         TabIndex        =   9
         Top             =   2085
         Width           =   765
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VENDAS"
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
      Left            =   510
      TabIndex        =   19
      Top             =   165
      Width           =   1560
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VENDAS"
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
      Left            =   540
      TabIndex        =   20
      Top             =   195
      Width           =   1560
   End
End
Attribute VB_Name = "frmSaidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddProd_Click()
    AdicionarProduto
End Sub

Private Sub cmdAdicionar_Click()
    Trava False
    Me.datVendas.Recordset.AddNew
    Me.dbcFuncionario.SetFocus
End Sub

Private Sub cmdCancelar_Click()
    cmdEditar_Click
    If Me.datVendas.EditMode = dbEditNone Then Exit Sub
    Me.datVendas.Recordset.CancelUpdate
    Trava True
End Sub

Private Sub cmdDeletar_Click()
    EXCLUIR_PRODUTOS
End Sub

Private Sub cmdEditar_Click()
    If Me.datVendas.Recordset.RecordCount = 0 Then Exit Sub
    Trava False
    Me.datVendas.Recordset.Edit
End Sub

Private Sub cmdExcluir_Click()
    Dim MSG As String
    MSG = "Deseja excluir ?"
    If Me.datVendas.Recordset.RecordCount = 0 Then Exit Sub
    If MsgBox(MSG, vbQuestion + vbYesNo, "Excluir") = vbNo Then Exit Sub
    Me.datVendas.Recordset.Delete
    If Me.datVendas.Recordset.RecordCount = 0 Then
        Me.datVendas.Refresh
    Else
        Me.datVendas.Recordset.MoveNext
        If Me.datVendas.Recordset.EOF Then
            Me.datVendas.Recordset.MoveLast
        End If
    End If
End Sub

Private Sub cmdSalvar_Click()
    Dim MSG As String
    MSG = "O cliente, o funcionário e a data da venda são obrigotórias."
    If Me.datVendas.EditMode = dbEditNone Then Exit Sub
    If Trim(Me.dbcFuncionario.Text) = "" Then
        MsgBox MSG, , "Crítica"
        Exit Sub
    End If
    If Trim(Me.dbcClientes.Text) = "" Then
        MsgBox MSG, , "Crítica"
        Exit Sub
    End If
    If Not IsDate(Me.txtDTVenda.Text) Then
        MsgBox MSG, , "Crítica"
        Exit Sub
    End If
    Me.datVendas.Recordset.Update
    Me.datVendas.Recordset.Bookmark = _
        Me.datVendas.Recordset.LastModified
    Trava True
End Sub

Private Sub DatProduto_Reposition()
    Me.txtPreco.Text = IIf(Not IsNull(Me.DatProduto.Recordset("PrecoUnitario")), Me.DatProduto.Recordset("PrecoUnitario"), 0)
    Me.txtQtd.Text = "0"
End Sub

Private Sub datVendas_Reposition()
    On Error Resume Next
    AtualizaGrid
    Me.datVendas.Caption = "Registro " & Me.datVendas.Recordset.AbsolutePosition + 1
End Sub

Private Sub Form_Activate()
        Trava True
        AtualizaGrid
        Me.txtPreco.Text = IIf(Not IsNull(Me.DatProduto.Recordset("PrecoUnitario")), Me.DatProduto.Recordset("PrecoUnitario"), 0)
        Me.txtQtd.Text = "0"
End Sub

Private Sub Form_Load()
    CaminhoBanco Me
End Sub

Private Sub Trava(Sim As Boolean)
    Me.dbcFuncionario.Locked = Sim
    Me.dbcClientes.Locked = Sim
    Me.txtDTVenda.Locked = Sim
    Me.datVendas.Visible = Sim
    Me.tabVendas.TabVisible(1) = Sim
End Sub

Private Sub AtualizaGrid()
    Dim Cod As Long
    If Me.DatProduto.Recordset.RecordCount = 0 Then Exit Sub
    If IsNull(Me.datVendas.Recordset("codigo")) Then Exit Sub
    Cod = CLng(Me.datVendas.Recordset("codigo"))
    Me.datFaturados.RecordSource = "SELECT * FROM tblHistoricoDeVenda WHERE CodVenda = " & Cod
End Sub

Private Sub AdicionarProduto()
    Dim MSG As String
    MSG = "Os dados não estão completos"
    If Trim(Me.dbcClientes.Text) = "" Then
        MsgBox MSG, , "Atenção"
        Exit Sub
    End If
    If Trim(Me.dbcFuncionario.Text) = "" Then
        MsgBox MSG, , "Atenção"
        Exit Sub
    End If
    If Not IsDate(Me.txtDTVenda.Text) Then
        MsgBox MSG, , "Atenção"
        Exit Sub
    End If
    If Not IsNumeric(Me.txtPreco.Text) Then
        MsgBox MSG, , "Atenção"
        Exit Sub
    End If
    If Not IsNumeric(Me.txtQtd.Text) Then
        MsgBox MSG, , "Atenção"
        Exit Sub
    End If
    If Me.DatProduto.Recordset.RecordCount = 0 Then
        MsgBox "Não existem produtos no extoque.", , "Atenção"
        Exit Sub
    End If
    If IsNull(Me.DatProduto.Recordset("Codigo")) Then
        MsgBox "Não existe produto selecionado.", , "Atenção"
        Exit Sub
    End If
    If Me.txtQtd.Text <= 0 Then
        MsgBox "A quantidade deve ser maior que 0.", , "Atenção"
        Exit Sub
    End If
    
    Dim RC As Recordset
    Set RC = Me.datFaturados.Recordset.Clone
    RC.AddNew
    RC("CodVenda") = CLng(Me.datVendas.Recordset("Codigo"))
    RC("Produto") = CLng(Me.DatProduto.Recordset("Codigo"))
    RC("Quantidade") = CLng(Me.txtQtd.Text)
    RC("ValorDaVenda") = CLng(Me.txtQtd.Text) * CLng(Me.txtPreco.Text)
    RC.Update
    Set RC = Nothing
    Me.DatProduto.Recordset.Edit
    Me.DatProduto.Recordset("QuantidadeEstoque") = _
        Me.DatProduto.Recordset("QuantidadeEstoque") - CLng(Me.txtQtd.Text)
    Me.DatProduto.Recordset.Update
    AtualizaGrid
    Me.datFaturados.Refresh
End Sub

Private Sub EXCLUIR_PRODUTOS()
    Do Until Me.datFaturados.Recordset.RecordCount = 0
        Me.datFaturados.Recordset.MoveFirst
        Me.datFaturados.Recordset.Delete
    Loop
    Me.datFaturados.Refresh
End Sub

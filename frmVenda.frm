VERSION 5.00
Begin VB.Form frmVenda 
   Caption         =   "V E N D A S"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   9210
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "&Adicionar"
      Height          =   330
      Left            =   3600
      TabIndex        =   5
      Top             =   3960
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   4695
      TabIndex        =   6
      Top             =   3960
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Excluir"
      Height          =   330
      Left            =   5730
      TabIndex        =   7
      Top             =   3960
      Width           =   1170
   End
   Begin VB.Data datVenda 
      Appearance      =   0  'Flat
      Caption         =   "Navega entre os registros."
      Connect         =   "Access"
      DatabaseName    =   "C:\Alunos\tarde\ESTOQUE\estoque.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H80000001&
      Height          =   345
      Left            =   180
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "tblHistoricoDeVenda"
      Top             =   3960
      Width           =   3225
   End
   Begin VB.TextBox txtCodigo 
      DataField       =   "Codigo"
      DataSource      =   "datVenda"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   330
      Width           =   1395
   End
   Begin VB.TextBox txtNome 
      DataField       =   "Nome"
      DataSource      =   "datVenda"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   870
      Width           =   4095
   End
   Begin VB.TextBox txtEndereco 
      DataField       =   "Endereco"
      DataSource      =   "datVenda"
      Height          =   375
      Left            =   1800
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1380
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      DataField       =   "Codigo"
      DataSource      =   "datVenda"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1920
      Width           =   1395
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nome"
      DataSource      =   "datVenda"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   2415
      Width           =   1845
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "PRODUTO.:"
      Height          =   195
      Left            =   780
      TabIndex        =   12
      Top             =   945
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "CODIGO.:"
      Height          =   195
      Left            =   945
      TabIndex        =   11
      Top             =   420
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "DATA DA VENDA.:"
      Height          =   195
      Left            =   285
      TabIndex        =   10
      Top             =   1995
      Width           =   1395
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "VALOR DA VENDA.:"
      Height          =   195
      Left            =   210
      TabIndex        =   9
      Top             =   2520
      Width           =   1500
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "QUANTIDADE.:"
      Height          =   195
      Left            =   525
      TabIndex        =   8
      Top             =   1470
      Width           =   1155
   End
End
Attribute VB_Name = "frmVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If Not Me.datVenda.Recordset.BOF And Not Me.datVenda.Recordset.EOF Then
       Me.datVenda.Recordset.Delete
       Me.datVenda.Recordset.MovePrevious
    If Me.datVenda.Recordset.BOF Then Me.datVenda.Recordset.MoveNext
    End If
End Sub

Private Sub Command2_Click()
    Me.datVenda.Recordset.Edit
    Me.datVenda.Recordset.Canc
End Sub

Private Sub Command3_Click()
On Error Resume Next
    Me.txtNome.SetFocus
    Me.datVenda.Recordset.AddNew
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

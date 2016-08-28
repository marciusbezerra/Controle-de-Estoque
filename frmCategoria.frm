VERSION 5.00
Begin VB.Form frmCategoria 
   Caption         =   "Cadastros"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4695
   ScaleWidth      =   6345
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Localização"
      Height          =   810
      Left            =   240
      TabIndex        =   12
      Top             =   2700
      Width           =   5865
      Begin VB.TextBox txtLocalizar 
         Height          =   405
         Left            =   960
         TabIndex        =   2
         Top             =   255
         Width           =   4755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   345
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdEditar 
      Appearance      =   0  'Flat
      Caption         =   "&Editar"
      Height          =   345
      Left            =   2550
      TabIndex        =   5
      Top             =   4170
      Width           =   1170
   End
   Begin VB.CommandButton cmdSalvar 
      Appearance      =   0  'Flat
      Caption         =   "&Salvar"
      Height          =   345
      Left            =   3750
      TabIndex        =   6
      Top             =   4170
      Width           =   1170
   End
   Begin VB.CommandButton cmdExcluir 
      Appearance      =   0  'Flat
      Caption         =   "E&xcluir"
      Height          =   345
      Left            =   1335
      TabIndex        =   4
      Top             =   4170
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4950
      TabIndex        =   7
      Top             =   4170
      Width           =   1170
   End
   Begin VB.CommandButton cmdAdicionar 
      Appearance      =   0  'Flat
      Caption         =   "&Adicionar"
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   4170
      Width           =   1170
   End
   Begin VB.TextBox txtComentarios 
      DataField       =   "Comentarios"
      DataSource      =   "datCategorias"
      Height          =   345
      Left            =   1245
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2085
      Width           =   4710
   End
   Begin VB.TextBox txtDescricao 
      DataField       =   "Descricao"
      DataSource      =   "datCategorias"
      Height          =   345
      Left            =   1245
      TabIndex        =   0
      Top             =   1560
      Width           =   4710
   End
   Begin VB.TextBox txtCodigo 
      DataField       =   "Codigo"
      DataSource      =   "datCategorias"
      Height          =   345
      HideSelection   =   0   'False
      Left            =   1245
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1035
      Width           =   1425
   End
   Begin VB.Data datCategorias 
      Appearance      =   0  'Flat
      Caption         =   "Navega entre os registros."
      Connect         =   "Access"
      DatabaseName    =   "C:\Alunos\tarde\ESTOQUE\estoque.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2865
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblCategorias"
      Top             =   3660
      Width           =   3225
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORIAS"
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
      Left            =   390
      TabIndex        =   15
      Top             =   255
      Width           =   2490
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORIAS"
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
      Left            =   420
      TabIndex        =   14
      Top             =   285
      Width           =   2490
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Comentário:"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   2145
      Width           =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   540
      TabIndex        =   10
      Top             =   1095
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
      Height          =   195
      Left            =   315
      TabIndex        =   9
      Top             =   1620
      Width           =   765
   End
End
Attribute VB_Name = "frmCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdicionar_Click()
    Trava False
    Me.datCategorias.Recordset.AddNew
    Me.txtDescricao.SetFocus
End Sub

Private Sub cmdCancelar_Click()
    cmdEditar_Click
    If Me.datCategorias.EditMode = dbEditNone Then Exit Sub
    Me.datCategorias.Recordset.CancelUpdate
    Trava True
End Sub

Private Sub cmdEditar_Click()
    If Me.datCategorias.Recordset.RecordCount = 0 Then Exit Sub
    Trava False
    Me.datCategorias.Recordset.Edit
End Sub

Private Sub cmdExcluir_Click()
    Dim MSG As String
    MSG = "Deseja excluir ?"
    If Me.datCategorias.Recordset.RecordCount = 0 Then Exit Sub
    If MsgBox(MSG, vbQuestion + vbYesNo, "Excluir") = vbNo Then Exit Sub
    Me.datCategorias.Recordset.Delete
    If Me.datCategorias.Recordset.RecordCount = 0 Then
        Me.datCategorias.Refresh
    Else
        Me.datCategorias.Recordset.MoveNext
        If Me.datCategorias.Recordset.EOF Then
            Me.datCategorias.Recordset.MoveLast
        End If
    End If
End Sub

Private Sub cmdSalvar_Click()
    If Me.datCategorias.EditMode = dbEditNone Then Exit Sub
    Me.datCategorias.Recordset.Update
    Me.datCategorias.Recordset.Bookmark = _
        Me.datCategorias.Recordset.LastModified
    Trava True
End Sub

Private Sub datCategorias_Reposition()
    On Error Resume Next
    Me.datCategorias.Caption = "Registro " & Me.datCategorias.Recordset.AbsolutePosition + 1
End Sub

Private Sub Form_Activate()
        Trava True
End Sub

Private Sub Form_Load()
    CaminhoBanco Me
End Sub



Private Sub txtLocalizar_Change()
    If Me.datCategorias.Recordset.RecordCount = 0 Then Exit Sub
   Me.datCategorias.Recordset.FindFirst "Descricao like '*" & Me.txtLocalizar.Text & "*'"
End Sub


Private Sub Trava(Sim As Boolean)
    Me.txtDescricao.Locked = Sim
    Me.txtComentarios.Locked = Sim
    Me.txtLocalizar.Locked = Not Sim
End Sub


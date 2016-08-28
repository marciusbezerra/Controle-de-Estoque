VERSION 5.00
Begin VB.Form frmFuncionarios 
   Caption         =   "Funcionários."
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4530
   ScaleWidth      =   6435
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Localização"
      Height          =   810
      Left            =   300
      TabIndex        =   14
      Top             =   2535
      Width           =   5865
      Begin VB.TextBox txtLocalizar 
         Height          =   345
         Left            =   960
         TabIndex        =   3
         Top             =   285
         Width           =   4755
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   345
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdAdicionar 
      Appearance      =   0  'Flat
      Caption         =   "&Adicionar"
      Height          =   345
      Left            =   195
      TabIndex        =   4
      Top             =   4005
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5025
      TabIndex        =   8
      Top             =   4005
      Width           =   1170
   End
   Begin VB.CommandButton cmdExcluir 
      Appearance      =   0  'Flat
      Caption         =   "E&xcluir"
      Height          =   345
      Left            =   1410
      TabIndex        =   5
      Top             =   4005
      Width           =   1170
   End
   Begin VB.CommandButton cmdSalvar 
      Appearance      =   0  'Flat
      Caption         =   "&Salvar"
      Height          =   345
      Left            =   3825
      TabIndex        =   7
      Top             =   4005
      Width           =   1170
   End
   Begin VB.CommandButton cmdEditar 
      Appearance      =   0  'Flat
      Caption         =   "&Editar"
      Height          =   345
      Left            =   2625
      TabIndex        =   6
      Top             =   4005
      Width           =   1170
   End
   Begin VB.TextBox txtDatadeInscricao 
      DataField       =   "DataDeEscricao"
      DataSource      =   "datFuncionarios"
      Height          =   330
      Left            =   4680
      TabIndex        =   2
      Top             =   1950
      Width           =   1485
   End
   Begin VB.Data datFuncionarios 
      Appearance      =   0  'Flat
      Caption         =   "Navega entre os registros."
      Connect         =   "Access"
      DatabaseName    =   "D:\Dados\PROG\VBASIC\Estoque\estoque.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2955
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblFuncionarios"
      Top             =   3480
      Width           =   3225
   End
   Begin VB.TextBox txtCodigo 
      DataField       =   "Codigo"
      DataSource      =   "datFuncionarios"
      Height          =   330
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1095
      Width           =   1470
   End
   Begin VB.TextBox txtNome 
      DataField       =   "NomeCompleto"
      DataSource      =   "datFuncionarios"
      Height          =   330
      Left            =   1200
      TabIndex        =   0
      Top             =   1530
      Width           =   4950
   End
   Begin VB.TextBox txtApelido 
      DataField       =   "Apelido"
      DataSource      =   "datFuncionarios"
      Height          =   330
      Left            =   1200
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1950
      Width           =   2295
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FUNCIONÁRIOS"
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
      Left            =   405
      TabIndex        =   16
      Top             =   240
      Width           =   2940
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Dt. Inscr.:"
      Height          =   195
      Left            =   3780
      TabIndex        =   13
      Top             =   2025
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   390
      TabIndex        =   12
      Top             =   1155
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
      Height          =   195
      Left            =   465
      TabIndex        =   11
      Top             =   1575
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Apelido:"
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   2010
      Width           =   570
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FUNCIONÁRIOS"
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
      TabIndex        =   17
      Top             =   270
      Width           =   2940
   End
End
Attribute VB_Name = "frmFuncionarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdicionar_Click()
    Trava False
    Me.datFuncionarios.Recordset.AddNew
    Me.txtNome.SetFocus
End Sub

Private Sub cmdCancelar_Click()
    cmdEditar_Click
    If Me.datFuncionarios.EditMode = dbEditNone Then Exit Sub
    Me.datFuncionarios.Recordset.CancelUpdate
    Trava True
End Sub

Private Sub cmdEditar_Click()
    If Me.datFuncionarios.Recordset.RecordCount = 0 Then Exit Sub
    Trava False
    Me.datFuncionarios.Recordset.Edit
End Sub

Private Sub cmdExcluir_Click()
    Dim MSG As String
    MSG = "Deseja excluir ?"
    If Me.datFuncionarios.Recordset.RecordCount = 0 Then Exit Sub
    If MsgBox(MSG, vbQuestion + vbYesNo, "Excluir") = vbNo Then Exit Sub
    Me.datFuncionarios.Recordset.Delete
    If Me.datFuncionarios.Recordset.RecordCount = 0 Then
        Me.datFuncionarios.Refresh
    Else
        Me.datFuncionarios.Recordset.MoveNext
        If Me.datFuncionarios.Recordset.EOF Then
            Me.datFuncionarios.Recordset.MoveLast
        End If
    End If
End Sub

Private Sub cmdSalvar_Click()
    If Me.datFuncionarios.EditMode = dbEditNone Then Exit Sub
    Me.datFuncionarios.Recordset.Update
    Me.datFuncionarios.Recordset.Bookmark = _
        Me.datFuncionarios.Recordset.LastModified
    Trava True
End Sub

Private Sub datFuncionarios_Reposition()
    On Error Resume Next
    Me.datFuncionarios.Caption = "Registro " & Me.datFuncionarios.Recordset.AbsolutePosition + 1
End Sub

Private Sub Form_Activate()
        Trava True
End Sub

Private Sub Form_Load()
    CaminhoBanco Me
End Sub



Private Sub txtLocalizar_Change()
   If Me.datFuncionarios.Recordset.RecordCount = 0 Then Exit Sub
   Me.datFuncionarios.Recordset.FindFirst "NomeCompleto like '*" & Me.txtLocalizar.Text & "*'"
End Sub


Private Sub Trava(Sim As Boolean)
    Me.txtNome.Locked = Sim
    Me.txtApelido.Locked = Sim
    Me.txtDatadeInscricao.Locked = Sim
    Me.txtLocalizar.Locked = Not Sim
End Sub




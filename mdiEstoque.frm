VERSION 5.00
Begin VB.MDIForm mdiEstoque 
   BackColor       =   &H8000000C&
   Caption         =   "Controle de Estoque."
   ClientHeight    =   5055
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Cadastro"
      Begin VB.Menu mnuCategoria 
         Caption         =   "Ca&tegoria"
      End
      Begin VB.Menu mnuClientes 
         Caption         =   "C&lientes"
      End
      Begin VB.Menu mnuProdutos 
         Caption         =   "&Produtos"
      End
      Begin VB.Menu mnuFuncionarios 
         Caption         =   "&Funcionarios"
      End
      Begin VB.Menu mnuFornecedores 
         Caption         =   "F&ornecedores"
      End
   End
   Begin VB.Menu mnuMovimento 
      Caption         =   "&Movimento"
      Begin VB.Menu mnuEntradas 
         Caption         =   "&Entradas "
      End
      Begin VB.Menu mnuVendas 
         Caption         =   "&Vendas no Varejo"
      End
      Begin VB.Menu mnuSaidas 
         Caption         =   "&Saidas"
      End
   End
   Begin VB.Menu mnuUtilitarios 
      Caption         =   "&Utilitários"
   End
   Begin VB.Menu mnuSair 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "mdiEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   
Private Sub mnuCategoria_Click()
   frmCategoria.Show
   frmCategoria.txtDescricao.SetFocus
End Sub
Private Sub mnuClientes_Click()
   frmClientes.Show
   frmClientes.txtNome.SetFocus
End Sub

Private Sub mnuEntradas_Click()
   frmEntradas.Show
   frmEntradas.txtProduto.SetFocus
End Sub

Private Sub mnuFornecedores_Click()
   frmFornecedores.Show
   frmFornecedores.txtRazaoSocial.SetFocus
End Sub

Private Sub mnuFuncionarios_Click()
   frmFuncionarios.Show
   frmFuncionarios.txtNome.SetFocus
End Sub

Private Sub mnuProdutos_Click()
   frmProdutos.Show
   frmProdutos.txtDescricao.SetFocus
End Sub

Private Sub mnuSaidas_Click()
   frmSaidas.Show
 End Sub

Private Sub mnuVendas_Click()
   frmVenda.Show
'   frmVenda.txtNome.SetFocus
End Sub

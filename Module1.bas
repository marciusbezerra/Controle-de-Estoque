Attribute VB_Name = "Module1"
Public marca As String
'1. Cria��o da tabela tblVendas
'     Codigo -> AutoNumera��o
'     CodFuncionario -> N�mero
'     CodCliente -> N�mero
'     DataDaVenda -> Data/Hora -> Valor Padr�o = Data()
'2. Modifica��o do desenho da tabela tblHistoricoDeVenda
'     Codigo -> AutoNumera��o
'     CodVenda -> N�mero
'     Produto -> N�mero
'     Quantidade -> N�mero
'     ValorDaVenda -> Moeda
'
'3. Cria��o dos relacionamentos de tabelas

Public Sub CaminhoBanco(Formul�rio As Form)
    Dim CaminhoProjeto As String
    Dim Controle As Control
    CaminhoProjeto = App.Path
    If Right(CaminhoProjeto, 1) <> "\" Then
        CaminhoProjeto = CaminhoProjeto & "\"
    End If
    CaminhoProjeto = CaminhoProjeto & "Estoque.mdb"
    For Each Controle In Formul�rio
        If TypeOf Controle Is Data Then
            Controle.DatabaseName = CaminhoProjeto
        End If
    Next
End Sub


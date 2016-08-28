Attribute VB_Name = "Module1"
Public marca As String
'1. Criação da tabela tblVendas
'     Codigo -> AutoNumeração
'     CodFuncionario -> Número
'     CodCliente -> Número
'     DataDaVenda -> Data/Hora -> Valor Padrão = Data()
'2. Modificação do desenho da tabela tblHistoricoDeVenda
'     Codigo -> AutoNumeração
'     CodVenda -> Número
'     Produto -> Número
'     Quantidade -> Número
'     ValorDaVenda -> Moeda
'
'3. Criação dos relacionamentos de tabelas

Public Sub CaminhoBanco(Formulário As Form)
    Dim CaminhoProjeto As String
    Dim Controle As Control
    CaminhoProjeto = App.Path
    If Right(CaminhoProjeto, 1) <> "\" Then
        CaminhoProjeto = CaminhoProjeto & "\"
    End If
    CaminhoProjeto = CaminhoProjeto & "Estoque.mdb"
    For Each Controle In Formulário
        If TypeOf Controle Is Data Then
            Controle.DatabaseName = CaminhoProjeto
        End If
    Next
End Sub


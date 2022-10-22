Attribute VB_Name = "Módulo1"
Sub Estrutura()
'Para declarar variável no VBA usamos o comando DIM

Dim Produto As String
Dim Preco As Double
Dim Desconto As Double
Dim Precofinal As Double

'Vamos utilizar a Caixa de Entrada, Inputbox para as variáveis

Produto = InputBox("Digite o Nome do produto", "Produto")
Preco = InputBox("Digite o Preço do produto", "Preço")
Desconto = InputBox("Digite o Desconto", "Desconto")
Precofinal = Preco - Preco * Desconto

Range("A1").Value = Produto
Range("A2").Value = Preco
Range("A3").Value = Desconto
Range("A4").Value = Precofinal

End Sub


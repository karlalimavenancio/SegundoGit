Attribute VB_Name = "M�dulo1"
Sub Constantes()
Const a1 As String = "A1"
Const a2 As String = "A2"

Dim Nome As String
Dim Numero As Integer

Nome = InputBox("Digite o seu nome")
Numero = InputBox("Digite um n�mero")

Range(a1).Value = Nome
If (Numero Mod 2 = 0) Then
Range(a2).Value = "Este n�mero � Par"
Else
Range(a2).Value = "Este n�mero � �mpar"
End If

End Sub

Sub MediaEscolar()

Const Media_Aprovacao As Double = 7

'Para notas maiores ou iguais a 7: Aprovado
'Para notas menores ou iguais a 4: Reprovado
'Para o restante: Recupera��o

Dim Nota As Double
Nota = InputBox("Digite a nota do aluno")

If (Nota > 10 Or Nota < 0) Then
MsgBox ("Nota Inv�lida")
Else

If (Nota >= Media_Aprovacao) Then
MsgBox ("Aprovado")
ElseIf (Nota <= 4) Then
MsgBox ("Reprovado")
Else
MsgBox ("Recupera��o")

End If
End If

End Sub


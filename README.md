# Visual Basic 6 Cheatsheet
 Um guia para os principais comandos da linguagem Visual Basic 6, para ser usado como uma forma de consulta rápida.
 
 

## TIPOS DE DADOS
```vb
 Dim age As Integer          ' Inteiro
 Dim coordinate As Long      ' Inteiro para números muito grandes
 Dim weight As Single        ' Float
 Dim pi As Double            ' Float com números muito grandes
 
 Dim name As String          ' String
 Dim birthDay As Date        ' Data (01/01/0100 up to 12/31/9999)
 Dim isRaining As Boolean    ' Bool
 Dim whataver As Variant     ' Guarda qualquer tipo de dado, é o tipo padrão de dado NÃO É RECOMENDADO USA-LA
 
```

## DECLARANDO VARIÁVEIS
```vb
 Dim name As String     ' A keyword Dim declara uma varíavel 
 Const COLOR As String  ' A keyword Const declara uma constante 
 
```

## OPERADORES ARITMÉTICOS
```vb
 Dim Num1, Num2 As Integer
 Dim result As Single
 
 Num1 = 5
 Num2 = 10
 
 result = Num1 +   Num2      ' Soma
 result = Num1 -   Num2      ' Subtração
 result = Num1 /   Num2      ' Divisão
 result = Num1 \   Num2      ' Divisão Inteira
 result = Num1 *   Num2      ' Multiplicação
 result = Num1 ^   Num2      ' Potenciação
 result = Num1 Mod Num2      ' Resto da divisão
 result = Num1 &   Num2      ' Concatenação
 
```

## OPERADORES RELACIONAIS
```vb
 Dim Num1, Num2 As Integer
 Dim result As Boolean
 
 Num1 = 5
 Num2 = 10
 
 result = Num1 >  Num2       ' Checa se Num1 é maior que Num2
 result = Num1 <  Num2       ' Checa se Num1 é menor que Num2
 result = Num1 >= Num2       ' Checa se Num1 é maior ou igual a Num2
 result = Num1 <= Num2       ' Checa se Num1 é menor ou igual a Num2
 result = Num1 <> Num2       ' Checa se Num1 não é igual ao Num2
 result = Num1 =  Num2       ' Checa se Num1 é igual ao Num2
 
```

## OPERADORES LÓGICOS
```vb
 Dim Num1, Num2 As Integer
 Dim result As Boolean

 Num1 = 5
 Num2 = 10
 
 result = Num1 > Num2 And Num1 > 0   ' Retorna True apenas se as duas operações forem verdadeiras
 result = Num1 > Num2 Or Num1  > 0   ' Retorna True apenas se uma das duas operações forem verdadeiras
 
```
 
## ESTRUTURAS CONDICIONAIS
```vb
 Dim Num1, Num2 As Integer
 Dim result As Boolean
 
 Num1 = 5
 Num2 = 10
 
 If (Num1 < Num2) then
  result = "minor"
 End If
 
 If (Num1 < Num2) then
  result = "minor"
 Else
  result = "greater"
 End If
 
 Select Case Num1
  Case 1
   result = "Num1 == 1"
  Case 2
   result = "Num1 == 2"
  Case 3
   result = "Num1 == 3" 
  Case 4
   result = "Num1 == 4"
 End Select   
  
```

## ESTRUTURAS DE REPETIÇÕES

## FUNCÕES

## SUB-ROTINAS

## CLASSES

## ATRIBUTO DE CLASSES

## METODOS DE CLASSES


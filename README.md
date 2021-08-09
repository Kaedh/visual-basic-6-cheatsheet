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
 Dim whataver As Variant     ' Guarda qualquer tipo de dado, é o tipo padrão de dado 
                             ' NÃO É RECOMENDADO USA-LA
 
```

## DECLARANDO VARIÁVEIS
```vb
 Dim name As String     ' A keyword Dim declara uma varíavel 
 Const COLOR As String  ' A keyword Const declara uma constante 
 Dim names(5) As String ' Cria um array de 5 elementos todos do tipo String
 
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
__________________________________________

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
```vb
 Dim number As Integer
 number = 1 

 Do While number <= 100 
  number = number + 1 
 Loop
 
 __________________________________________

 number = 1

 While number <=100
  number = number + 1 
 Wend
 
 __________________________________________
 
 Dim number As Long
 number = 0
 
 Do  
  number = number + 1 
 Loop While number < 201
 
 __________________________________________
 
 Dim number As Long
 number = 0
 
 Do Until number > 1000 
  number = number + 1 
  Print number 
 Loop
 
 __________________________________________
 
 Dim x As Integer
 
 For x = 1 To 50 
  Print x 
 Next 
 
```

## FUNCÕES
```vb
 Funções em visual basic SEMPRE RETORNAM ALGUM VALOR para o controle

 Public Function soma(a As Integer, b As Integer) As Integer
 
  soma = a + b             ' chamando o nome da funcão e atribuindo  
                           ' um valor, é a forma de fazer retorno da funcão
                   
 End Function
 
 Private Sub Command1_Click()
 
 Debug.Print (soma(5, 9))  ' Para invocar uma função basta escrever o nome
                           ' o nome da função seguida de ( ) dentro será
                           ' passado os argumentos caso tenha

 End Sub
 
```


## SUB-ROTINAS
```vb
 Sub-rotinas em Visual Basic funcionam parecido com funções,
 Mas elas não precisam retornar nenhum valor para o controle.

 Private Sub HelloWorld()
     Debug.Print "Hello World"
 End Sub
 
 Private Sub Command1_Click()
    HelloWorld       ' Escreve "Hello World" no console
 End Sub
```

## CLASSES
```vb
 Option Explicit
     Private cName As String                                  ' Declarando os atributos da classe
     Private cAge As Integer

 Public Sub Class_Initialize()
     cName = ""                                               ' Função que é executada quando o 
     cAge = 0                                                 ' objeto é montado
 End Sub

 Public Property Get name()
     name = cName                                             ' Método get do atributo cName
 End Property

 Public Property Let setNewName(newName As String)
  cName = newName                                             ' Método set do atributo CName
 End Property

 Public Property Get age()
     age = cAge                                               ' Método get do atributo cAge
 End Property

 Public Property Let setNewAge(newAge As Integer)
  cAge = newAge                                               ' Método set do atributo cAge
 End Property

 Public Sub Apresentação()
     MsgBox "Nome : " & cName & " Idade: " & cAge             ' Método que faz o uso dos 
 End Sub                                                      ' atributos  interno da classe
 
```


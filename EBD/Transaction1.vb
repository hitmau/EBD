' Transaction1
Imports System.EnterpriseServices
Imports System

' O atributo Transaction torna sua classe preparada para transações.  Sua classe pode configurar o tipo de transação do seu 
' objeto para um dos seguintes:
' 
' Required
' Required New
' Supported
' Not Supported
' Disabled

<Transaction(TransactionOption.Supported)> _
Public Class Transaction1
    Inherits ServicedComponent

    ' Implemente os métodos da sua classe aqui.
    '
    ' Componentes com transação utilizam o objeto ContextUtil para notificar o caller se eles finalizaram
    ' com sucesso ou não.  Se a transação sucedeu, o método deve ativar
    ' ContextUtil.SetComplete.  Caso contrário, o método deve ativar
    ' ContextUtil.SetAbort.
    '
    ' Public Sub MySub()
    '    Try
    '        " Código do que deve ocorrer na transação aqui.
    '        " Sem erros.  Declara que a transação pode terminar com SetComplete
    '        ContextUtil.SetComplete()
    '    Catch ex As Exception
    '        " Uma exceção foi lançada durante a transação.  
    '        " A transação não pode finalizar e SetAbort foi chamado.
    '        contextutil.SetAbort()
    '    End Try
    ' End Sub

    ' Ao invés de ajustar explicitamente o estado de ContextUtil , métodos em uma classe com transação podem utilizar 
    ' o atributo AutoComplete.  Se o método retornar com sucesso, SetComplete será chamado.
    ' Se o métodos lança uma exceção, SetAbort será chamado.
    ' 
    ' <AutoComplete()> Public Sub MyMethod()
    ' End Sub

End Class

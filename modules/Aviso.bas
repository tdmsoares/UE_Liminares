Option Explicit

Function Cadastro(Formul�rio As String) As Boolean
Dim Aviso As VbMsgBoxResult

Aviso = MsgBox("Este acesso � exclusivo para altera��o e/ou cadastro. Deseja continuar?", vbExclamation + vbYesNo)

If (Aviso = vbYes) Then
'
'Abre o formul�rio Senha
'DoCmd.openForm("")
'
'Abre o formul�rio CadastroAlunos
DoCmd.OpenForm (Formul�rio)
Cadastro = True
Exit Function

Else:
Cadastro = False
Exit Function
End If

End Function
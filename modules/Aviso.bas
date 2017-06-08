Option Explicit

Function Cadastro(Formulário As String) As Boolean
Dim Aviso As VbMsgBoxResult

Aviso = MsgBox("Este acesso é exclusivo para alteração e/ou cadastro. Deseja continuar?", vbExclamation + vbYesNo)

If (Aviso = vbYes) Then
'
'Abre o formulário Senha
'DoCmd.openForm("")
'
'Abre o formulário CadastroAlunos
DoCmd.OpenForm (Formulário)
Cadastro = True
Exit Function

Else:
Cadastro = False
Exit Function
End If

End Function
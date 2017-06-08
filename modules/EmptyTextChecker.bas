Option Explicit

Function isEmptyText(ByVal Text) As Boolean
    '
    'Checks whether the Text is null or empty
    Dim isEmpty As Boolean
    isEmpty = False
    '
    If (IsNull(Text)) Then
        isEmpty = True
    ElseIf (Text = "") Then
        isEmpty = True
    ElseIf (Trim(Text)) = "" Then
        isEmpty = True
    End If
    '
    isEmptyText = isEmpty
End Function

Sub AvisoPreenchimentoIncompleto(Campo, TextoMensagem, Requerido As Boolean)
'
'Lida caso haja Campo não preenchido
If (Requerido = False) Then
    Dim corrigir As VbMsgBoxResult
    corrigir = MsgBox(TextoMensagem & vbCrLf & "Deseja corrigir?", vbQuestion + vbYesNo, "Cadastro Incompleto")
    If (corrigir = vbYes) Then
        Campo.SetFocus
        End
    End If
Else:
    MsgBox TextoMensagem & vbCrLf & "Corriga e Tente Novamente", vbExclamation, "Campo Requerido não Preenchido"
    Campo.SetFocus
    End
End If

End Sub

Function VerificarCadastroIncompleto(Campo) As Boolean '
'
'Verifica os campos em Branco
    If (isEmptyText(Campo)) Then
        VerificarCadastroIncompleto = True
    Else
        VerificarCadastroIncompleto = False
    End If
End Function
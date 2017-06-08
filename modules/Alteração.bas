Option Explicit

Sub Dados(Formulário As String, Campo As String, DadosAtuais, Id)
'
'Registra qualquer Alteração
Dim Db As DAO.Database
Dim rg As Recordset

Set Db = CurrentDb
Set rg = Db.OpenRecordset("Alteração", dbOpenDynaset)

With rg
    .AddNew
    !Formulário = Formulário
    !Campo = Campo
    ![Dado Atual] = DadosAtuais
    !IdRegistro = Id
    !Data = DateTime.Now
    !Computador = User.GetCurrentPC
    ![Nome Usuário] = User.GetCurrentUser
    .Update
    .Bookmark = rg.LastModified
End With

rg.Close

End Sub
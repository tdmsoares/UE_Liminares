Option Explicit

Sub Dados(Formul�rio As String, Campo As String, DadosAtuais, Id)
'
'Registra qualquer Altera��o
Dim Db As DAO.Database
Dim rg As Recordset

Set Db = CurrentDb
Set rg = Db.OpenRecordset("Altera��o", dbOpenDynaset)

With rg
    .AddNew
    !Formul�rio = Formul�rio
    !Campo = Campo
    ![Dado Atual] = DadosAtuais
    !IdRegistro = Id
    !Data = DateTime.Now
    !Computador = User.GetCurrentPC
    ![Nome Usu�rio] = User.GetCurrentUser
    .Update
    .Bookmark = rg.LastModified
End With

rg.Close

End Sub
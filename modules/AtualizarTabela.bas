Option Explicit

Public Sub AtualizarLiminaresPendentes()
On Error GoTo Errado
'
'Atribui as tabelas
Dim Db As DAO.Database
Dim rgCrian�a As Recordset
Dim rgAcompanhamento As Recordset
Dim rgLiminaresPendentes As Recordset
'
Dim C�digoCrian�a
Dim IdCrian�a
Dim Pendente As Boolean
Dim TipoOcorr�ncia
Dim DataOcorr�ncia
'
Set Db = CurrentDb
Set rgCrian�a = Db.OpenRecordset("Crian�a", dbOpenDynaset)
Set rgAcompanhamento = Db.OpenRecordset("Acompanhamento", dbOpenDynaset)
Set rgLiminaresPendentes = Db.OpenRecordset("Liminares_Pendentes", dbOpenDynaset)
'
'Verifica se a tabela Liminares_Pendentes n�o est� vazia
'Se tiver registros estes ser�o excluidos
If (rgLiminaresPendentes.RecordCount <> 0) Then
    '
    'Excluindo registros (Por c�digo sql - consulta Exclus�o)
    DoCmd.SetWarnings False
    Dim sql As String
    sql = "DELETE Liminares_Pendentes.*" & _
            "FROM Liminares_Pendentes;"
    '
    DoCmd.RunSQL (sql)
    DoCmd.SetWarnings False
End If
DoCmd.SetWarnings True
'
'Coloca a tabela crian�a com foco no primeiro registro
If (rgCrian�a.BOF = False) Then
    rgCrian�a.MoveFirst
End If
'
'Faz a busca se h� registros de ocorr�ncias da liminar em Acompanhamento da crian�a referida na tabela Crian�a
Do
    '
    Pendente = False
    C�digoCrian�a = rgCrian�a!C�digo
    '
    'Coloca a tabela Acompanhamento com foco no primeiro registro
    If (rgAcompanhamento.BOF = False) Then
        rgAcompanhamento.MoveFirst
    End If
    '
    'Depois verifica se est� no Aguardo
    If (rgCrian�a!Status = "Aguardo") Then
        Pendente = True
        Do
            IdCrian�a = rgAcompanhamento!IdCrian�a
            '
            'Primeiro busca o registro da crian�a na tabela Acompanhamento
            If (C�digoCrian�a = IdCrian�a) Then
                    If (rgAcompanhamento!Ocorr�ncia <> "Disponibilizada") Then
                        Pendente = False
                    Else:
                        TipoOcorr�ncia = rgAcompanhamento!Ocorr�ncia
                        DataOcorr�ncia = rgAcompanhamento!Data
                    End If

            End If
            rgAcompanhamento.MoveNext
        Loop Until rgAcompanhamento.EOF
            If Pendente = True Then
                If DateDiff("d", DataOcorr�ncia, DateTime.Now) > 30 Then
                    With rgLiminaresPendentes
                        .AddNew
                        !IdCrian�a = C�digoCrian�a
                        !Nome = rgCrian�a!Nome
                        !Nascimento = rgCrian�a!Nascimento
                        !Status = rgCrian�a!Status
                        !Data = DataOcorr�ncia
                        !Ocorr�ncia = TipoOcorr�ncia
                        .Update
                        .Bookmark = .LastModified
                    End With
                End If
            End If
    End If
    rgCrian�a.MoveNext
Loop Until rgCrian�a.EOF


Errado:
Resume Next
End Sub

Sub AtualizarPesquisaR�pida()
On Error GoTo Errado
'
'Atribui as tabelas
Dim Db As DAO.Database
Dim rgPesquisaR�pidaCrian�a As Recordset
Dim rgConsultaPesquisaR�pidaCrian�a As Recordset
'
Set Db = CurrentDb
Set rgPesquisaR�pidaCrian�a = Db.OpenRecordset("PesquisaR�pidaCrian�a", dbOpenDynaset)
Set rgConsultaPesquisaR�pidaCrian�a = Db.OpenRecordset("Crian�a", dbOpenDynaset)
'
'Verifica se j� foi preenchida a tavela PesquisaR�pidaAlunos
'Caso sim excluir os registros para deixar a tabela vazia antes de acrescentar os registros
If (rgPesquisaR�pidaCrian�a.RecordCount <> 0) Then
    '
    'Excluindo registros (Por c�digo sql - consulta Exclus�o)
    DoCmd.SetWarnings False
    Dim sql As String
    sql = "DELETE PesquisaR�pidaCrian�a.*" & _
            "FROM PesquisaR�pidaCrian�a;"
    '
    DoCmd.RunSQL (sql)
    DoCmd.SetWarnings False
End If
DoCmd.SetWarnings True
'
'Coloca a tabela CadastroAlunos com foco no primeiro registro
If (rgConsultaPesquisaR�pidaCrian�a.BOF = False) Then
    rgConsultaPesquisaR�pidaCrian�a.MoveFirst
End If
'
'Faz a busca se h� registros de ocorr�ncias da liminar em Acompanhamento da crian�a referida na tabela Crian�a
Do
    
    With rgPesquisaR�pidaCrian�a
        .AddNew
        !C�digo = rgConsultaPesquisaR�pidaCrian�a!C�digo
        !Nome = rgConsultaPesquisaR�pidaCrian�a!Nome
        !Nascimento = rgConsultaPesquisaR�pidaCrian�a!Nascimento
        !Status = rgConsultaPesquisaR�pidaCrian�a!Status
        !Processo = rgConsultaPesquisaR�pidaCrian�a!Processo
        .Update
        .Bookmark = .LastModified
    End With
    rgConsultaPesquisaR�pidaCrian�a.MoveNext
Loop Until rgConsultaPesquisaR�pidaCrian�a.EOF

DoCmd.Requery
Errado:
Resume Next
End Sub
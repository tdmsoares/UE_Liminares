Option Explicit

Public Sub AtualizarLiminaresPendentes()
On Error GoTo Errado
'
'Atribui as tabelas
Dim Db As DAO.Database
Dim rgCriança As Recordset
Dim rgAcompanhamento As Recordset
Dim rgLiminaresPendentes As Recordset
'
Dim CódigoCriança
Dim IdCriança
Dim Pendente As Boolean
Dim TipoOcorrência
Dim DataOcorrência
'
Set Db = CurrentDb
Set rgCriança = Db.OpenRecordset("Criança", dbOpenDynaset)
Set rgAcompanhamento = Db.OpenRecordset("Acompanhamento", dbOpenDynaset)
Set rgLiminaresPendentes = Db.OpenRecordset("Liminares_Pendentes", dbOpenDynaset)
'
'Verifica se a tabela Liminares_Pendentes não está vazia
'Se tiver registros estes serão excluidos
If (rgLiminaresPendentes.RecordCount <> 0) Then
    '
    'Excluindo registros (Por código sql - consulta Exclusão)
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
'Coloca a tabela criança com foco no primeiro registro
If (rgCriança.BOF = False) Then
    rgCriança.MoveFirst
End If
'
'Faz a busca se há registros de ocorrências da liminar em Acompanhamento da criança referida na tabela Criança
Do
    '
    Pendente = False
    CódigoCriança = rgCriança!Código
    '
    'Coloca a tabela Acompanhamento com foco no primeiro registro
    If (rgAcompanhamento.BOF = False) Then
        rgAcompanhamento.MoveFirst
    End If
    '
    'Depois verifica se está no Aguardo
    If (rgCriança!Status = "Aguardo") Then
        Pendente = True
        Do
            IdCriança = rgAcompanhamento!IdCriança
            '
            'Primeiro busca o registro da criança na tabela Acompanhamento
            If (CódigoCriança = IdCriança) Then
                    If (rgAcompanhamento!Ocorrência <> "Disponibilizada") Then
                        Pendente = False
                    Else:
                        TipoOcorrência = rgAcompanhamento!Ocorrência
                        DataOcorrência = rgAcompanhamento!Data
                    End If

            End If
            rgAcompanhamento.MoveNext
        Loop Until rgAcompanhamento.EOF
            If Pendente = True Then
                If DateDiff("d", DataOcorrência, DateTime.Now) > 30 Then
                    With rgLiminaresPendentes
                        .AddNew
                        !IdCriança = CódigoCriança
                        !Nome = rgCriança!Nome
                        !Nascimento = rgCriança!Nascimento
                        !Status = rgCriança!Status
                        !Data = DataOcorrência
                        !Ocorrência = TipoOcorrência
                        .Update
                        .Bookmark = .LastModified
                    End With
                End If
            End If
    End If
    rgCriança.MoveNext
Loop Until rgCriança.EOF


Errado:
Resume Next
End Sub

Sub AtualizarPesquisaRápida()
On Error GoTo Errado
'
'Atribui as tabelas
Dim Db As DAO.Database
Dim rgPesquisaRápidaCriança As Recordset
Dim rgConsultaPesquisaRápidaCriança As Recordset
'
Set Db = CurrentDb
Set rgPesquisaRápidaCriança = Db.OpenRecordset("PesquisaRápidaCriança", dbOpenDynaset)
Set rgConsultaPesquisaRápidaCriança = Db.OpenRecordset("Criança", dbOpenDynaset)
'
'Verifica se já foi preenchida a tavela PesquisaRápidaAlunos
'Caso sim excluir os registros para deixar a tabela vazia antes de acrescentar os registros
If (rgPesquisaRápidaCriança.RecordCount <> 0) Then
    '
    'Excluindo registros (Por código sql - consulta Exclusão)
    DoCmd.SetWarnings False
    Dim sql As String
    sql = "DELETE PesquisaRápidaCriança.*" & _
            "FROM PesquisaRápidaCriança;"
    '
    DoCmd.RunSQL (sql)
    DoCmd.SetWarnings False
End If
DoCmd.SetWarnings True
'
'Coloca a tabela CadastroAlunos com foco no primeiro registro
If (rgConsultaPesquisaRápidaCriança.BOF = False) Then
    rgConsultaPesquisaRápidaCriança.MoveFirst
End If
'
'Faz a busca se há registros de ocorrências da liminar em Acompanhamento da criança referida na tabela Criança
Do
    
    With rgPesquisaRápidaCriança
        .AddNew
        !Código = rgConsultaPesquisaRápidaCriança!Código
        !Nome = rgConsultaPesquisaRápidaCriança!Nome
        !Nascimento = rgConsultaPesquisaRápidaCriança!Nascimento
        !Status = rgConsultaPesquisaRápidaCriança!Status
        !Processo = rgConsultaPesquisaRápidaCriança!Processo
        .Update
        .Bookmark = .LastModified
    End With
    rgConsultaPesquisaRápidaCriança.MoveNext
Loop Until rgConsultaPesquisaRápidaCriança.EOF

DoCmd.Requery
Errado:
Resume Next
End Sub